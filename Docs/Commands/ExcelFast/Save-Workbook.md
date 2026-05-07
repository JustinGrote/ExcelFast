---
document type: cmdlet
external help file: PowerShell.dll-Help.xml
HelpUri: ''
Locale: en-US
Module Name: ExcelFast
ms.date: 05/07/2026
PlatyPS schema version: 2024-05-01
title: Save-Workbook
---

# Save-Workbook

## SYNOPSIS

Saves an Excel workbook to a file.

## SYNTAX

### __AllParameterSets

```
Save-Workbook [-Workbook] <IXLWorkbook> [[-Destination] <string>] [-Force] [<CommonParameters>]
```

## ALIASES

This cmdlet has the following aliases:
- `svwb`

## DESCRIPTION

The Save-Workbook cmdlet saves an XLWorkbook object to an Excel file. If a destination is not specified, the workbook is saved to its original file location. Multiple workbooks can be saved by piping them to the cmdlet. The -Force parameter can be used to overwrite existing files or create directories.

## EXAMPLES

### Example 1: Save workbook to its original location

```powershell
$workbook = Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
Save-Workbook -Workbook $workbook
```

Saves the workbook back to its original file location.

### Example 2: Save workbook to a new location

```powershell
$workbook = Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
Save-Workbook -Workbook $workbook -Destination 'C:\Data\Backup.xlsx'
```

Saves the workbook to a new file path.

### Example 3: Overwrite existing file

```powershell
$workbook = Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
Save-Workbook -Workbook $workbook -Destination 'C:\Data\Export.xlsx' -Force
```

Uses the -Force parameter to overwrite an existing file.

### Example 4: Use pipeline input

```powershell
Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx' | Save-Workbook -Destination 'C:\Data\Copy.xlsx'
```

Gets the workbook and pipes it to Save-Workbook to save to a new location.

### Example 5: Create directory and save

```powershell
$workbook = Get-Workbook -Path 'C:\Data\Spreadsheet.xlsx'
Save-Workbook -Workbook $workbook -Destination 'C:\Backup\Archive\Export.xlsx' -Force
```

Uses -Force to create the directory path if it doesn't exist.

### Example 6: Save multiple workbooks

```powershell
@('C:\Data\File1.xlsx', 'C:\Data\File2.xlsx') | Get-Workbook | Save-Workbook
```

Saves multiple workbooks back to their original locations.

## PARAMETERS

### -Destination

Specifies the file path where the workbook will be saved. If not specified, the workbook will be saved to its original file location.

```yaml
Type: System.String
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
HelpMessage: 'Destination where the Excel file will be saved. If not specified, the workbook will be saved to its current location.'
```

### -Force

Overwrites the destination file if it exists. Can also be used to create directory paths that do not exist.

```yaml
Type: System.Management.Automation.SwitchParameter
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
HelpMessage: 'If specified, overwrites the file if it exists.'
```

### -Workbook

Specifies the XLWorkbook object to save. Obtain this from Get-Workbook.

```yaml
Type: ClosedXML.Excel.IXLWorkbook
DefaultValue: $null
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: 'The workbook to save.'
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### ClosedXML.Excel.IXLWorkbook

Accepts XLWorkbook objects from Get-Workbook.

## OUTPUTS

None. This cmdlet does not produce output objects. Use -Verbose to see status messages.

## NOTES

- Cannot use -Destination parameter with multiple workbooks; save individually instead
- If directory does not exist, use -Force to create the directory path
- If file exists, use -Force to overwrite it
- The workbook must be properly opened with Get-Workbook before saving
- Use -Verbose to see detailed status messages

## RELATED LINKS

- [Get-Workbook](Get-Workbook.md)
- [Import-Workbook](Import-Workbook.md)
- [Export-Workbook](Export-Workbook.md)

