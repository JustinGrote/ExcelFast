using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;

using MiniExcelLibs;

using static ExcelFast.Constants;

using FilePath = System.IO.Path;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Import, CmdletDefaultName)]
[OutputType(typeof(PSObject))]
[OutputType(typeof(IEnumerable<dynamic>), ParameterSetName = [nameof(Raw)])]
public class ImportCommand : BaseCmdlet
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

	[Parameter(
		HelpMessage = "Specify the ending cell for data import (e.g., 'A1', 'B2'). This is only used when NoHeaders is set to true."
	)]
	public string EndCell { get; set; } = string.Empty;

	[Parameter(
		HelpMessage = "Return the result as a raw dynamic enumerable without PSObject wrapping. Use only for advanced performance use cases.",
		ParameterSetName = nameof(Raw)
	)]
	private static SwitchParameter Raw { get; set; } = false;

	readonly HashSet<ICollection<string>> columnSets = [];
	private const string ImportedPSTypeName = "ExcelFast.ImportedWorkbook";

	protected override void ProcessRecord()
	{
		foreach (string pathItem in Path)
		{
			string providerPath = GetUnresolvedProviderPathFromPSPath(pathItem);
			Debug($"Importing Workbook: {providerPath}");

			if (!File.Exists(providerPath))
			{
				Error(new FileNotFoundException($"Excel file not found: {providerPath}"),
					"Check the file path and try again.",
					"FileNotFound",
					providerPath
				);
				continue;
			}

			string fileExtension = FilePath.GetExtension(providerPath).ToLowerInvariant();
			if (!AcceptedExtensions.Contains(fileExtension))
			{
				Error(
					new ArgumentException($"Unsupported file type '{fileExtension}' for '{providerPath}'.", "Path"),
					$"Use one of the supported file types: {string.Join(", ", AcceptedExtensions)}",
					"UnsupportedFileType",
					providerPath
				);
				continue;
			}

			IEnumerable<dynamic> rows = [];
			try
			{
				if (string.IsNullOrWhiteSpace(SheetName))
				{
					Debug($"No sheet name provided. Importing the first sheet from '{providerPath}'.");
				}
				else if (!MiniExcel.GetSheetNames(providerPath).Contains(SheetName, StringComparer.OrdinalIgnoreCase))
				{
					Error(
						new ArgumentException($"Sheet '{SheetName}' does not exist in the '{providerPath}' workbook."),
						"Check the sheet name and try again.",
						"InvalidSheetName",
						SheetName
					);
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
					Warning($"Sheet '{SheetName}' in '{providerPath}' has different columns than previously imported sheets. The resultant object output may be different and not displayed correctly.");
					columnSets.Add(columns);
				}

				rows = string.IsNullOrEmpty(EndCell)
					? MiniExcel.Query(
						providerPath,
						useHeaderRow: !NoHeaders.IsPresent,
						sheetName: SheetName,
						startCell: StartCell
					)
					: MiniExcel.QueryRange(
						providerPath,
						!NoHeaders.IsPresent,
						SheetName,
						startCell: StartCell,
						endCell: EndCell
					);


			}
			catch (NotSupportedException ex)
			{
				if (!ex.Message.Contains("Stream cannot know the file type"))
				{
					throw;
				}
				Error(
					new ArgumentException($"{providerPath} is not a valid Excel file."),
					"UnsupportedFileType",
					providerPath
				);
			}

			if (Raw.IsPresent)
			{
				// Return the raw enumerable as-is so the consumer can stream/transform using their preferred method.
				WriteObject(rows, false);
				continue;
			}

			foreach (IDictionary<string, object> row in rows)
			{
				// Create PSObject directly from properties to avoid intermediate allocations
				if (row.Count == 0)
				{
					Debug($"Row in '{providerPath}' is empty. Skipping.");
					continue;
				}

				PSObject psObject = new(row.Count);
				psObject.TypeNames.Insert(0, ImportedPSTypeName);

				// BUG: WriteObject(data, true) enumerates everything before pipelining so we cant use it here. Need to file an issue.
				foreach (KeyValuePair<string, object> property in row)
				{
					psObject.Properties.Add(new PSNoteProperty(property.Key, property.Value));
				}
				WriteObject(psObject);
			}
		}
	}
}
