---
document type: cmdlet
external help file: ExcelFast-Help.xml
HelpUri: ''
Locale: en-US
Module Name: ExcelFast
ms.date: 05/07/2026
PlatyPS schema version: 2024-05-01
title: Import-Workbook
---

# Import-Workbook

## SYNOPSIS

Imports data from an Excel workbook, worksheet, or range into PowerShell objects.

## SYNTAX

### Path

```
Import-Workbook [-Path] <string[]> [[-SheetName] <string[]>] [-NoHeaders] [-StartCell <string>]
 [-EndCell <string>] [-Raw] [-IncludeEmptyRows] [<CommonParameters>]
```

### Workbook

```
Import-Workbook [-Workbook] <IXLWorkbook> [[-SheetName] <string[]>] [-NoHeaders]
 [-StartCell <string>] [-EndCell <string>] [-Raw] [-IncludeEmptyRows] [<CommonParameters>]
```

### Range

```
Import-Workbook [-Range] <IXLRangeBase> [[-SheetName] <string[]>] [-NoHeaders] [-StartCell <string>]
 [-EndCell <string>] [-Raw] [-IncludeEmptyRows] [<CommonParameters>]
```

## ALIASES

This cmdlet has the following aliases:
- `iwb`


## DESCRIPTION

The Import-Workbook cmdlet imports data from Excel files, workbooks, worksheets, or ranges and converts them into PowerShell objects. It supports three input modes: file paths, XLWorkbook objects, or XLRangeBase objects (from worksheets or tables). By default, the first row is treated as column headers to create object properties. Use -NoHeaders to treat all rows as data.

## EXAMPLES

### Example 1: Import data from a file

```powershell
Import-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
```

Imports data from the first sheet of the Excel file. The first row is treated as column headers.

### Example 2: Import data using pipeline

```powershell
Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx' | Import-Workbook
```

Gets the workbook and pipes it to Import-Workbook to import the data.

### Example 3: Import specific sheet

```powershell
Import-Workbook -Path 'C:\Data\Spreadsheet.xlsx' -SheetName 'Sales'
```

Imports data from the sheet named 'Sales' instead of the first sheet.

### Example 4: Import data without headers

```powershell
Import-Workbook -Path 'C:\Data\Spreadsheet.xlsx' -NoHeaders
```

Imports all rows as data without treating the first row as headers. Properties will be named Column1, Column2, etc.

### Example 5: Import from a specific cell range

```powershell
Import-Workbook -Path 'C:\Data\Spreadsheet.xlsx' -StartCell 'B2' -EndCell 'D10'
```

Imports data from the range B2:D10, useful when headers are not in the first row.

### Example 6: Import from workbook object

```powershell
$workbook = Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
$worksheet = $workbook.Worksheets | Where-Object Name -eq 'Sheet1'
$worksheet | Import-Workbook
```

Imports data by piping a worksheet object to the cmdlet.

### Example 7: Import data with empty rows included

```powershell
Import-Workbook -Path 'C:\Data\Spreadsheet.xlsx' -IncludeEmptyRows
```

Includes empty rows in the output. By default, empty rows are skipped.

## PARAMETERS

### -EndCell

Specifies the ending cell for data import (e.g., 'D10'). This is only used in conjunction with -StartCell or when -NoHeaders is set to true.
Specify the ending cell for data import (e.g., 'A1', 'B2'). This is only used when NoHeaders is set to true.

```yaml
Type: String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -IncludeEmptyRows

Includes empty rows in the output. By default, empty rows are skipped.
Include empty rows in the output. By default, empty rows are skipped.

```yaml
Type: SwitchParameter
DefaultValue: false
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -NoHeaders

Specifies that the first row should not be used as column headers. All rows will be treated as data.
Do not use the first row as column headers.

```yaml
Type: SwitchParameter
DefaultValue: false
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Path

Specifies the path to the Excel file to import. Can accept multiple file paths for batch import.
Path to the Excel file to import.

```yaml
Type: String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: Path
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Range

Specifies a range object to import. Accepts worksheet ranges, table ranges, or workbook ranges obtained from Get-Workbook.
Range to import. Accepts Table Ranges, Worksheet Ranges, or Workbook Ranges. Get using Get-Workbook, select the appropriate Worksheet, and then select the appropriate Range from the Ranges property.

```yaml
Type: IXLRangeBase
DefaultValue: $null
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: Range
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Raw

Returns the result as a raw dynamic enumerable without PSObject wrapping. Use only for advanced performance use cases.
Return the result as a raw dynamic enumerable without PSObject wrapping. Use only for advanced performance use cases.

```yaml
Type: SwitchParameter
DefaultValue: false
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -SheetName

Specifies the name(s) of the sheet(s) to import. If not specified, imports from the first sheet.
Names of sheet(s) to import. If not specified, imports the first sheet.

```yaml
Type: String[]
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 1
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -StartCell

Specifies the starting cell for data import (e.g., 'A1', 'B2'). Useful when data doesn't start at A1.
Specify the starting cell for data import (e.g., 'A1', 'B2').

```yaml
Type: String
DefaultValue: A1
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Workbook

Specifies the XLWorkbook object to import from. Get this from Get-Workbook.
Workbook object to import. Get using Get-Workbook.

```yaml
Type: IXLWorkbook
DefaultValue: $null
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: Workbook
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### String

Accepts file paths as input.

### String

Accepts an array of file paths for batch import.

### IXLWorkbook

Accepts XLWorkbook objects from Get-Workbook.

### IXLRangeBase

Accepts range objects from worksheets or tables.

### String[]

{{ Fill in the Description }}

## OUTPUTS

### PSObject

Returns PSObject instances representing each row of data, with properties named from the header row.

### Object

{{ Fill in the Description }}

## NOTES

- Supports .xlsx and .csv file formats
- The first row is treated as headers by default; use -NoHeaders to disable this
- Empty rows are skipped by default; use -IncludeEmptyRows to include them
- Cannot use -Destination parameter with multiple workbooks; save individually instead

## RELATED LINKS

- [Get-Workbook](Get-Workbook.md)
- [Save-Workbook](Save-Workbook.md)
- [Export-Workbook](Export-Workbook.md)
