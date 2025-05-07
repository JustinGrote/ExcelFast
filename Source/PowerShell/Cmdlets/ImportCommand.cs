using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;

using MiniExcelLibs;

using static ExcelFast.Constants;

using FilePath = System.IO.Path;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Import, CmdletDefaultName)]
[OutputType(typeof(IEnumerable<dynamic>))]
public class ImportCommand : PSCmdlet
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

	[Parameter(
		HelpMessage = "Do not use the first row as column headers."
	)]
	public SwitchParameter NoHeaders { get; set; }

	[Parameter(
		HelpMessage = "Specify the starting cell for data import (e.g., 'A1', 'B2')."
	)]
	public string StartCell { get; set; } = "A1";

	// Used in logging
	string name => MyInvocation.MyCommand.Name;

	//Used to detect if columns have changed during import so we can warn
	readonly HashSet<ICollection<string>> columnSets = [];

	// PSTypeName to be added to all objects
	private const string ImportedPSTypeName = "ExcelFast.ImportedWorkbook";

	protected override void ProcessRecord()
	{
		foreach (string pathItem in Path)
		{
			string providerPath = GetUnresolvedProviderPathFromPSPath(pathItem);
			WriteDebug($"{name}: Importing Workbook: {providerPath}");

			if (!File.Exists(providerPath))
			{
				WriteError(new ErrorRecord(
						new FileNotFoundException($"Excel file not found: {providerPath}"),
						"FileNotFound",
						ErrorCategory.ObjectNotFound,
						providerPath
				));
				continue;
			}

			string fileExtension = FilePath.GetExtension(providerPath).ToLowerInvariant();
			if (!AcceptedExtensions.Contains(fileExtension))
			{
				WriteError(new ErrorRecord(
					new ArgumentException($"Unsupported file type '{fileExtension}' for '{providerPath}'. Supported file types: {string.Join(',', AcceptedExtensions)}", "Path"),
					"UnsupportedFileType",
					ErrorCategory.InvalidArgument,
					providerPath
				));
				continue;
			}

			IEnumerable<dynamic> rows = [];
			try
			{
				if (string.IsNullOrWhiteSpace(SheetName))
				{
					WriteDebug($"{name}: No sheet name provided. Importing the first sheet from '{providerPath}'.");
				}
				else if (!MiniExcel.GetSheetNames(providerPath).Contains(SheetName, StringComparer.OrdinalIgnoreCase))
				{
					WriteError(new ErrorRecord(
							new ArgumentException($"Sheet '{SheetName}' does not exist in the '{providerPath}' workbook."),
							"InvalidSheetName",
							ErrorCategory.InvalidArgument,
							SheetName
					));
					continue;
				}

				ICollection<string> columns = MiniExcel.GetColumns(
					providerPath,
					!NoHeaders.IsPresent,
					SheetName,
					startCell: StartCell
				);

				if (!columnSets.Any())
				{
					columnSets.Add(columns);
				}
				else if (!columnSets.Any(c => c.SequenceEqual(columns)))
				{
					WriteWarning($"Sheet '{SheetName}' in '{providerPath}' has different columns than previously imported sheets. The resultant object output may be different and not displayed correctly.");
					columnSets.Add(columns);
				}

				rows = MiniExcel.Query(
					providerPath,
					useHeaderRow: !NoHeaders.IsPresent,
					sheetName: SheetName,
					startCell: StartCell
				);
			}
			catch (NotSupportedException ex)
			{
				if (!ex.Message.Contains("Stream cannot know the file type"))
				{
					throw;
				}
				WriteError(new ErrorRecord(
						new ArgumentException($"{providerPath} is not a valid Excel file."),
						"UnsupportedFileType",
						ErrorCategory.InvalidData,
						providerPath
				));
			}

			// NOTE: WriteObject(data, true) enumerates everything before pipelining. Need to file an issue.
			foreach (IDictionary<string, object> row in rows)
			{
				// Create PSObject directly from properties to avoid intermediate allocations
				if (row.Count == 0)
				{
					WriteDebug($"Row in '{providerPath}' is empty. Skipping.");
					continue;
				}

				PSObject psObject = new(row.Count);
				psObject.TypeNames.Insert(0, ImportedPSTypeName);

				foreach (KeyValuePair<string, object> property in row)
				{
					psObject.Properties.Add(new PSNoteProperty(property.Key, property.Value));
				}
				WriteObject(psObject);
			}
		}
	}
}
