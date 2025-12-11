# OfficeMetaClean  
自动删除常见office文件中的所有属性信息。   
Remove all property information from Office files.

```
CleanMeta - WPS/Office 文档元数据清理工具

用法: 
  1. cleanmeta.exe [参数] <文件 或 文件夹>
  2. 多选拖放office文件或目录到主程序

参数说明: 
  -h         显示帮助
  -b         处理前在同目录备份原文件
  -l         按天归集留存日志
 
支持的格式:
  doc/docx/docm/wps/xls/xlsx/xlsm/et/ppt/pptx/pptm/dps

注意:
  1. 自动清除office文件中包含的所有属性信息；
  2. 处理doc/wps/xls/et/ppt/dps等文件需要本机安装WPS/Office；

示例：
  cleanmeta.exe D:\test.doc E:\test2.et
  cleanmeta.exe -b -l D:\docs\test.doc
  cleanmeta.exe -l D:\folder
```
