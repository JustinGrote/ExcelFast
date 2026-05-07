
namespace ExcelFast.PowerShell.Cmdlets;

using ClosedXML.Excel;

using static ExcelFast.SystemConstants;

[Cmdlet(VerbsCommon.Get, CmdletDefaultName)]
[OutputType(typeof(XLWorkbook))]
[Alias("gwb", "Open-Workbook", "owb")]
public class GetCommand : BaseCmdlet
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
		if (Path is null || Path.Length == 0)
		{
			return;
		}

		foreach (string pathItem in Path)
		{
			string resolvedPath;
			try
			{
				resolvedPath = ResolveWorkbookPath(pathItem, out string? tempPath);
			}
			catch (Exception ex)
			{
				Error(
					ex,
					"Check the URL and verify the remote file is reachable.",
					"RemoteFileDownloadFailed",
					pathItem
				);
				continue;
			}

			try
			{
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

				IXLWorkbook workbook = new XLWorkbook(resolvedPath);
				WriteObject(workbook);
			}
			catch (IOException ex) when (ex.HResult == ERROR_SHARING_VIOLATION)
			{
				Error(
					ex,
					"Ensure the workbook is not open in Excel or locked by another process, then try again.",
					"WorkbookFileLocked",
					resolvedPath
				);
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
