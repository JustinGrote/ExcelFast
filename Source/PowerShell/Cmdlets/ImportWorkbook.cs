using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;
using ClosedXML.Excel;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Import, "ExcelFastWorkbook")]
[OutputType(typeof(XLWorkbook))]
public class ImportExcelFastWorkbookCommand : PSCmdlet
{
	[Parameter(
			Mandatory = true,
			Position = 0,
			ValueFromPipeline = true,
			ValueFromPipelineByPropertyName = true,
			HelpMessage = "Path to the Excel file to import as a workbook."
	)]
	[ValidateNotNullOrEmpty]
	[NotNull]
	public string[]? Path { get; set; }

	protected override void ProcessRecord()
	{
		foreach (string pathItem in Path)
		{
			string resolvedPath = GetUnresolvedProviderPathFromPSPath(pathItem);

			if (!File.Exists(resolvedPath))
			{
				WriteError(new ErrorRecord(
						new FileNotFoundException($"Excel file not found: {resolvedPath}"),
						"FileNotFound",
						ErrorCategory.ObjectNotFound,
						resolvedPath
				));
				continue;
			}

			try
			{
				XLWorkbook workbook = new(resolvedPath);
				WriteObject(workbook);
			}
			catch (Exception ex)
			{
				WriteError(new ErrorRecord(
						ex,
						"ImportExcelWorkbookError",
						ErrorCategory.ReadError,
						resolvedPath
				));
			}
		}
	}
}
