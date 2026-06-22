---
document type: cmdlet
external help file: ExcelFast-Help.xml
HelpUri: ''
Locale: en-US
Module Name: ExcelFast
ms.date: 05/07/2026
PlatyPS schema version: 2024-05-01
title: Export-Workbook
---

# Export-Workbook

## SYNOPSIS

Exports PowerShell objects to an Excel workbook.

## SYNTAX

### __AllParameterSets

```
Export-Workbook [-Destination] <string> [-InputObject] <PSObject[]> [-SheetName <string>] [-Force]
 [<CommonParameters>]
```

## ALIASES

This cmdlet has the following aliases:
- `exwb`

## DESCRIPTION

The Export-Workbook cmdlet exports PowerShell objects to an Excel workbook. Objects are converted to rows in the Excel file, with object properties becoming column headers. The cmdlet supports exporting to a single sheet or multiple sheets. If the destination file exists and -Force is not specified, an error will be raised.

## EXAMPLES

### Example 1: Export objects to Excel

```powershell
$data = @(
    @{ Name = 'Item1'; Value = 100 },
    @{ Name = 'Item2'; Value = 200 }
)
Export-Workbook -Destination 'C:\Data\Export.xlsx' -InputObject $data
```

Exports an array of custom objects to a new Excel file.

### Example 2: Export pipeline data

```powershell
Get-Process | Select-Object Name, CPU, Memory | Export-Workbook -Destination 'C:\Data\Processes.xlsx' -Force
```

Exports process data from Get-Process to an Excel file, using -Force to overwrite if the file exists.

### Example 3: Export to specific sheet

```powershell
$users = @(
    @{ Username = 'user1'; Email = 'user1@example.com' },
    @{ Username = 'user2'; Email = 'user2@example.com' }
)
Export-Workbook -Destination 'C:\Data\Users.xlsx' -InputObject $users -SheetName 'Users'
```

Exports data to a named sheet instead of the default 'Sheet1'.

### Example 4: Export multiple data sets

```powershell
$sales = Get-Content -Path 'C:\Data\sales.json' | ConvertFrom-Json
$users = Import-Csv -Path 'C:\Data\users.csv'
Export-Workbook -Destination 'C:\Data\Report.xlsx' -InputObject @($sales, $users)
```

Exports multiple data sets to separate sheets in the workbook.

### Example 5: Export cmdlet output

```powershell
Get-ChildItem -Path 'C:\Data' -File | Export-Workbook -Destination 'C:\Data\FileList.xlsx' -Force
```

Exports file system objects to Excel with metadata.

### Example 6: Apply a table style

```powershell
$data = @(
    @{ Name = 'Item1'; Value = 100 },
    @{ Name = 'Item2'; Value = 200 }
)
Export-Workbook -Destination 'C:\Data\Export.xlsx' -InputObject $data -TableStyle 'TableStyleMedium2'
```

Exports the data and applies the requested table style to the resulting table.

### Example 7: Apply a table name

```powershell
$data = @(
    @{ Name = 'Item1'; Value = 100 },
    @{ Name = 'Item2'; Value = 200 }
)
Export-Workbook -Destination 'C:\Data\Export.xlsx' -InputObject $data -TableName 'SalesData'
```

Exports the data and assigns the specified name to the resulting table.

## PARAMETERS

### -Destination

Specifies the file path where the Excel workbook will be created. If the file already exists, use -Force to overwrite it.

```yaml
Type: String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: 'Path to the Excel file to export to.'
```

### -Force

Overwrites the destination file if it already exists, and creates necessary directory paths.

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
HelpMessage: 'Forces overwriting of the destination file if it already exists.'
```

### -InputObject

Specifies the PowerShell objects to export to the Excel file. Object properties are exported as column headers, and each object becomes a row.

```yaml
Type: PSObject[]
DefaultValue: $null
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 1
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: 'Objects to export to the Excel file.'
```

### -SheetName

Specifies the name of the sheet in the Excel workbook. If not specified, uses 'Sheet1'.

```yaml
Type: String
DefaultValue: 'Sheet1'
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
HelpMessage: 'Name of the sheet to export to. If not specified, exports to ''Sheet1''.'
```

### -TableStyle

Specifies the ClosedXML table style to apply to the exported table. Use tab completion to discover supported styles.

```yaml
Type: String
DefaultValue: 'TableStyleMedium2'
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
HelpMessage: 'Apply the specified ClosedXML table style to the exported table.'
```

### -TableName

Specifies the name of the exported table.

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
HelpMessage: 'Specify the name of the exported table.'
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### String

You can pipe file paths as the -Destination parameter.

### PSObject

You can pipe PowerShell objects to be exported to the -InputObject parameter.

## OUTPUTS

None. This cmdlet does not produce output objects. Use -Verbose to see status messages.

## NOTES

- Supported file formats: .xlsx and .csv
- Object properties become Excel column headers
- Each object becomes a row in the Excel sheet
- If no objects are provided, a warning message is displayed
- Use -Force to overwrite existing files or create directory paths
- When exporting multiple object arrays, each array is placed in a separate sheet
- Use -Verbose to see detailed export information

## RELATED LINKS

- [Get-Workbook](Get-Workbook.md)
- [Import-Workbook](Import-Workbook.md)
- [Save-Workbook](Save-Workbook.md)

