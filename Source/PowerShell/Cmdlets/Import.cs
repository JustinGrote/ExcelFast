using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;
using MiniExcelLibs;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Import, "ExcelFast")]
[OutputType(typeof(IEnumerable<dynamic>))]
public class ImportExcelFastCommand : PSCmdlet
{
	[Parameter(
			Mandatory = true,
			Position = 0,
			ValueFromPipeline = true,
			ValueFromPipelineByPropertyName = true,
			HelpMessage = "Path to the Excel file to import."
	)]
	[ValidateNotNullOrEmpty]
	[NotNull]
	public string[]? Path { get; set; }

	[Parameter(
			Position = 1,
			HelpMessage = "Names of sheets to import. If not specified, imports the first sheet."
	)]
	public string? SheetName { get; set; }

	const string name = "Import-ExcelFast";

	//Used to detect if columns have changed during import so we can warn
	HashSet<ICollection<string>> columnSets = [];

	protected override void ProcessRecord()
	{
		foreach (string pathItem in Path)
		{
			string resolvedPath = GetUnresolvedProviderPathFromPSPath(pathItem);
			WriteDebug($"{name}: Importing Workbook: {resolvedPath}");

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

			if (string.IsNullOrWhiteSpace(SheetName))
			{
				WriteDebug($"{name}: No sheet name provided. Importing the first sheet from '{resolvedPath}'.");
			}
			else
			{
				if (!MiniExcel.GetSheetNames(resolvedPath).Contains(SheetName, StringComparer.OrdinalIgnoreCase))
				{
					WriteError(new ErrorRecord(
							new ArgumentException($"Sheet '{SheetName}' does not exist in the '{resolvedPath}' workbook."),
							"InvalidSheetName",
							ErrorCategory.InvalidArgument,
							SheetName
					));
					continue;
				}
			}

			ICollection<string> columns = MiniExcel.GetColumns(resolvedPath, true, SheetName);
			if (!columnSets.Any())
			{
				columnSets.Add(columns);
			}
			else if (!columnSets.Any(c => c.SequenceEqual(columns)))
			{
				WriteWarning($"Sheet '{SheetName}' in '{resolvedPath}' has different columns than previously imported sheets. The resultant object output may be different and not displayed correctly.");
				columnSets.Add(columns);
			}

			IEnumerable<dynamic> rows = MiniExcel.Query(resolvedPath, true, SheetName);
			// NOTE: WriteObject(data, true) enumerates everything before pipelining. Need to file an issue.
			foreach (dynamic row in rows)
			{
				WriteObject(row);
			}
		}
	}
}
