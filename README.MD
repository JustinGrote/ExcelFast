# ExcelFast

ExcelFast is a PowerShell module designed to quickly import, export, and modify Excel and CSV workbooks.

**This is an alpha work-in-progress!**

## Quick Start
```powershell
#Import a worksheet
$data = Import-Workbook path/to/excel.xlsx

#Import only a few rows in a worksheet, note this is very fast expecially in a large workbook
$data = Import-Workbook path/to/excel.xlsx | Select -First 3

#Open a workbook for editing
$worksheet = Open-Workbook path/to/excel.xlsx
$worksheet.<tab> #Note there is full intellisense in vscode and excel

#Export data
$data | Export-Workbook path/to/newexcel.xlsx