---
document type: cmdlet
external help file: ExcelFast.dll-Help.xml
HelpUri: ''
Locale: en-US
Module Name: ExcelFast
ms.date: 05/07/2026
PlatyPS schema version: 2024-05-01
title: Get-Workbook
---

# Get-Workbook

## SYNOPSIS

Opens an Excel workbook for manipulation and analysis.

## SYNTAX

### __AllParameterSets

```
Get-Workbook [-Path] <string[]> [<CommonParameters>]
```

## ALIASES

This cmdlet has the following aliases:
- `gwb`
- `owb`
- `Open-Workbook`

## DESCRIPTION

The Get-Workbook cmdlet opens an Excel workbook file and returns an XLWorkbook object that can be used with other ExcelFast commands. It supports local file paths, as well as remote URLs for files stored online.

## EXAMPLES

### Example 1: Open a local Excel file

```powershell
Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
```

Opens the Excel file at the specified path and returns the workbook object.

### Example 2: Open multiple Excel files

```powershell
Get-Workbook -Path 'C:\Data\File1.xlsx', 'C:\Data\File2.xlsx'
```

Opens multiple Excel files and returns an array of workbook objects.

### Example 3: Use pipeline input with alias

```powershell
'C:\Data\Spreadsheet.xlsx' | gwb
```

Uses the short alias `gwb` and passes the file path through the pipeline.

### Example 4: Open workbook and access worksheet

```powershell
$workbook = Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
$worksheet = $workbook.Worksheets | Where-Object Name -eq 'Sheet1'
```

Opens a workbook and accesses a specific worksheet by name.

## PARAMETERS

### -Path

Specifies the path to the Excel file to open. Accepts local file paths or remote URLs. Supports multiple files for batch processing.

```yaml
Type: System.String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: 'Path to the Excel file to import as a workbook.'
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String

You can pipe one or more file paths as strings to this cmdlet.

### System.String[]

You can pipe an array of file paths to this cmdlet.

## OUTPUTS

### ClosedXML.Excel.XLWorkbook

The cmdlet returns one or more XLWorkbook objects representing the opened Excel files.

## NOTES

- If the workbook is open in Excel or locked by another process, an error will be returned.
- The workbook must be a valid Excel file (.xlsx or .csv format).
- Remote files are downloaded to a temporary location and opened from there.

## RELATED LINKS

- [Import-Workbook](Import-Workbook.md)
- [Save-Workbook](Save-Workbook.md)
- [Export-Workbook](Export-Workbook.md)

