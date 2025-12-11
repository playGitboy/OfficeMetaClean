package main

import (
    "archive/zip"
    "fmt"
    "io"
    "os"
    "flag"
    "path/filepath"
    "strings"
    "sync"
    "time"

    "github.com/go-ole/go-ole"
    "github.com/go-ole/go-ole/oleutil"
)

var officeExts = []string{
    ".doc", ".docx", ".docm", ".wps",
    ".et", ".xlsx", ".xls", ".xlsm",
    ".pps", ".ppt", ".pptx", ".pptm", ".dps",
}

var (
    enableBackup bool
    enableLog    bool
    logFile      *os.File
    logMutex     sync.Mutex
)

func main() {
    showHelp := flag.Bool("h", false, "help")
    flag.BoolVar(&enableBackup, "b", false, "backup")
    flag.BoolVar(&enableLog, "l", false, "log")
    flag.Parse()

    // 无路径参数 且 没有要求备份或日志 → 显示帮助
    if *showHelp || (len(flag.Args()) == 0 && !enableBackup && !enableLog) {
        exeDir, _ := filepath.Abs(filepath.Dir(os.Args[0]))
        helpPath := filepath.Join(exeDir, "help.txt")

        data, err := os.ReadFile(helpPath)
        if err == nil {
            os.Stdout.Write(data)
        }
        return
    }

    var paths []string
    for _, arg := range flag.Args() {
        absPath, err := filepath.Abs(arg)
        if err == nil {
            paths = append(paths, absPath)
        }
    }

    if enableLog {
        initLog(paths[0])
        defer logFile.Close()
    }

    // 收集文件
    var files []string
    for _, path := range paths {
        info, err := os.Stat(path)
        if err != nil {
            logPrintf("无法访问: %s, %v", path, err)
            continue
        }

        if info.IsDir() {
            filepath.Walk(path, func(p string, info os.FileInfo, err error) error {
                if !info.IsDir() && isOfficeFile(p) {
                    absPath, err := filepath.Abs(p)
                    if err == nil {
                        files = append(files, absPath)
                    }
                }
                return nil
            })
        } else if isOfficeFile(path) {
            files = append(files, path)
        }
    }

    if len(files) == 0 {
        logPrintf("未找到支持的文件")
        return
    }

    // 备份
    if enableBackup {
        for _, f := range files {
            err := backupFile(f)
            if err != nil {
                logPrintf("备份失败: %s, %v", f, err)
            } else {
                logPrintf("备份成功: %s", f)
            }
        }
    }

    // 初始化 COM
    ole.CoInitialize(0)
    defer ole.CoUninitialize()

    var converted []string
    for _, f := range files {
        logPrintf("处理文件: %s", f)

        cf, err := convertOldFile(f)
        if err != nil {
            logPrintf("转换失败: %s, %v", f, err)
            continue
        }

        waitFileReady(cf, 15)
        converted = append(converted, cf)
    }

    var wg sync.WaitGroup
    for _, f := range converted {
        wg.Add(1)
        go func(file string) {
            defer wg.Done()
            err := removePropertiesWithRetry(file, 3)
            if err != nil {
                logPrintf("删除属性失败: %s, %v", file, err)
            } else {
                logPrintf("删除属性成功: %s", file)
            }
        }(f)
    }
    wg.Wait()

    logPrintf("所有文件处理完成")
}

func isOfficeFile(fileName string) bool {
    ext := strings.ToLower(filepath.Ext(fileName))
    for _, e := range officeExts {
        if e == ext {
            return true
        }
    }
    return false
}

func initLog(basePath string) {
    dir := filepath.Join(filepath.Dir(basePath), "log")
    os.MkdirAll(dir, 0755)
    logFileName := filepath.Join(dir, time.Now().Format("20060102")+".log")
    f, err := os.OpenFile(logFileName, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
    if err != nil {
        return
    }
    logFile = f
}

func logPrintf(format string, args ...interface{}) {
    if !enableLog || logFile == nil {
        return
    }
    msg := fmt.Sprintf(format, args...)
    logMutex.Lock()
    defer logMutex.Unlock()
    logFile.WriteString(time.Now().Format("15:04:05 ") + msg + "\r\n")
}

func backupFile(filePath string) error {
    dir := filepath.Dir(filePath)
    base := filepath.Base(filePath)
    backupPath := filepath.Join(dir, base+".bak")
    src, err := os.Open(filePath)
    if err != nil {
        return err
    }
    defer src.Close()
    dst, err := os.Create(backupPath)
    if err != nil {
        return err
    }
    defer dst.Close()
    _, err = io.Copy(dst, src)
    return err
}

func convertOldFile(filePath string) (string, error) {
    ext := strings.ToLower(filepath.Ext(filePath))
    var newFile string

    switch ext {
    case ".doc", ".wps":
        newFile = strings.TrimSuffix(filePath, ext) + ".docx"
        if err := convertWordOrWPS(filePath, newFile, ext); err != nil {
            return "", err
        }
    case ".xls", ".et":
        newFile = strings.TrimSuffix(filePath, ext) + ".xlsx"
        if err := convertExcelOrET(filePath, newFile, ext); err != nil {
            return "", err
        }
    case ".ppt", ".dps":
        newFile = strings.TrimSuffix(filePath, ext) + ".pptx"
        if err := convertPowerPointOrDPS(filePath, newFile, ext); err != nil {
            return "", err
        }
    default:
        newFile = filePath
    }
    time.Sleep(1 * time.Second)
    return newFile, nil
}

func waitFileReady(path string, timeout int) {
    for i := 0; i < timeout; i++ {
        f, err := os.OpenFile(path, os.O_RDWR, 0644)
        if err == nil {
            f.Close()
            return
        }
        time.Sleep(1 * time.Second)
    }
    logPrintf("文件 %s 超时未准备好", path)
}

func convertWordOrWPS(src, dst, ext string) error {
    var progID string
    if ext == ".doc" {
        progID = "Word.Application"
    } else {
        progID = "KWPS.Application"
    }

    appObj, err := oleutil.CreateObject(progID)
    if err != nil {
        return fmt.Errorf("启动 %s COM 失败: %v", progID, err)
    }
    defer appObj.Release()
    app, _ := appObj.QueryInterface(ole.IID_IDispatch)
    defer app.Release()
    oleutil.PutProperty(app, "Visible", false)

    docs := oleutil.MustGetProperty(app, "Documents").ToIDispatch()
    defer docs.Release()

    absSrc, _ := filepath.Abs(src)
    absDst, _ := filepath.Abs(dst)
    doc := oleutil.MustCallMethod(docs, "Open", absSrc,
        false, false, false).ToIDispatch()
    defer doc.Release()

    // 注意AI或网上代码用“16”都是错误的，后面必须用“12”否则某些旧版WPS另存docx实际还是doc/wps格式
    _, err = oleutil.CallMethod(doc, "SaveAs2", absDst, 12)
    if err != nil {
        return err
    }

    oleutil.CallMethod(doc, "Close")
    oleutil.CallMethod(app, "Quit")
    time.Sleep(2 * time.Second)
    return nil
}

func convertExcelOrET(src, dst, ext string) error {
    var progID string
    if ext == ".xls" {
        progID = "Excel.Application"
    } else {
        progID = "ket.Application"
    }

    appObj, err := oleutil.CreateObject(progID)
    if err != nil {
        return fmt.Errorf("启动 %s COM 失败: %v", progID, err)
    }
    defer appObj.Release()
    app, _ := appObj.QueryInterface(ole.IID_IDispatch)
    defer app.Release()
    oleutil.PutProperty(app, "Visible", false)

    wbs := oleutil.MustGetProperty(app, "Workbooks").ToIDispatch()
    defer wbs.Release()
    absSrc, _ := filepath.Abs(src)
    absDst, _ := filepath.Abs(dst)
    wb := oleutil.MustCallMethod(wbs, "Open", absSrc).ToIDispatch()
    defer wb.Release()

    _, err = oleutil.CallMethod(wb, "SaveAs", absDst, 51)
    if err != nil {
        return err
    }

    oleutil.CallMethod(wb, "Close", false)
    oleutil.CallMethod(app, "Quit")
    time.Sleep(2 * time.Second)
    return nil
}

func convertPowerPointOrDPS(src, dst, ext string) error {
    var progID string
    if ext == ".ppt" {
        progID = "PowerPoint.Application"
    } else {
        progID = "dps.Application"
    }

    appObj, err := oleutil.CreateObject(progID)
    if err != nil {
        return fmt.Errorf("启动 %s COM 失败: %v", progID, err)
    }
    defer appObj.Release()
    app, _ := appObj.QueryInterface(ole.IID_IDispatch)
    defer app.Release()
    oleutil.PutProperty(app, "Visible", true)

    pres := oleutil.MustGetProperty(app, "Presentations").ToIDispatch()
    defer pres.Release()
    absSrc, _ := filepath.Abs(src)
    absDst, _ := filepath.Abs(dst)
    ppt := oleutil.MustCallMethod(pres, "Open", absSrc, false, false, false).ToIDispatch()
    defer ppt.Release()

    _, err = oleutil.CallMethod(ppt, "SaveAs", absDst, 24)
    if err != nil {
        return err
    }

    oleutil.CallMethod(ppt, "Close")
    oleutil.CallMethod(app, "Quit")
    time.Sleep(2 * time.Second)
    return nil
}

func removePropertiesWithRetry(filePath string, retry int) error {
    var err error
    if isZipFile(filePath) {
        for i := 0; i < retry; i++ {
            err = removeProperties(filePath)
            if err == nil {
                return nil
            }
            logPrintf("删除属性失败，重试 %d: %v", i+1, err)
            time.Sleep(1 * time.Second)
        }
    } else {
        err = fmt.Errorf("警告: 文件不是OOXML格式，请确认文件格式！")
    }
    return err
}

func isZipFile(file string) bool {
    f, err := os.Open(file)
    if err != nil { return false }
    defer f.Close()

    header := make([]byte, 4)
    if _, err := f.Read(header); err != nil {
        return false
    }
    return header[0] == 0x50 && header[1] == 0x4B
}

func removeProperties(filePath string) error {
    tmpDir := filePath + "_tmp"
    os.MkdirAll(tmpDir, 0755)

    r, err := zip.OpenReader(filePath)
    if err != nil {
        return err
    }
    defer r.Close()

    for _, f := range r.File {
        if strings.HasPrefix(f.Name, "docProps/") || strings.HasPrefix(f.Name, "customXml/") {
            continue
        }

        destPath := filepath.Join(tmpDir, f.Name)
        if f.FileInfo().IsDir() {
            os.MkdirAll(destPath, 0755)
            continue
        }

        os.MkdirAll(filepath.Dir(destPath), 0755)
        rc, err := f.Open()
        if err != nil {
            return err
        }

        outFile, err := os.Create(destPath)
        if err != nil {
            rc.Close()
            return err
        }

        _, err = io.Copy(outFile, rc)
        rc.Close()
        outFile.Close()
        if err != nil {
            return err
        }
    }

    err = zipDir(tmpDir, filePath)
    if err != nil {
        return err
    }

    os.RemoveAll(tmpDir)
    return nil
}

func zipDir(source, target string) error {
    outFile, err := os.Create(target)
    if err != nil {
        return err
    }
    defer outFile.Close()

    zw := zip.NewWriter(outFile)
    defer zw.Close()

    return filepath.Walk(source, func(path string, info os.FileInfo, err error) error {
        if info.IsDir() {
            return nil
        }

        relPath, err := filepath.Rel(source, path)
        if err != nil {
            return err
        }

        f, err := zw.Create(relPath)
        if err != nil {
            return err
        }

        srcFile, err := os.Open(path)
        if err != nil {
            return err
        }
        defer srcFile.Close()

        _, err = io.Copy(f, srcFile)
        return err
    })
}
