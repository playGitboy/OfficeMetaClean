SET GOOS=windows
SET GOARCH=386
"C:\Program Files\Go1.19\bin\go.exe" mod init cleanmeta
"C:\Program Files\Go1.19\bin\go.exe" mod tidy
"C:\Program Files\Go1.19\bin\go.exe" build -ldflags "-H=windowsgui" -o cleanmeta.exe cleanmeta.go
