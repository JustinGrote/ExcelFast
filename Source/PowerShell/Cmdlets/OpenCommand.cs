
namespace ExcelFast.PowerShell.Cmdlets;

using ClosedXML.Excel;

[Cmdlet(VerbsCommon.Open, CmdletDefaultName)]
[OutputType(typeof(XLWorkbook))]
[Alias("owb")]
public class OpenCommand : BaseCmdlet
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
				Error(
					new FileNotFoundException($"Excel file not found: {resolvedPath}"),
					"Verify the file path and try again.",
					"FileNotFound",
					resolvedPath
				);
				continue;
			}

			try
			{
				XLWorkbook workbook = new(resolvedPath);
				WriteObject(workbook);
			}
			catch (Exception ex)
			{
				Error(
					ex,
					"Check if the file is a valid Excel file and is not corrupted.",
					"ImportExcelWorkbookError",
					resolvedPath
				);
			}
		}
	}
}
