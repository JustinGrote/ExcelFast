namespace ExcelFast.PowerShell.Cmdlets;

using MiniExcelLibs;

using FilePath = Path;

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

	[Parameter(
		HelpMessage = "Include empty rows in the output. By default, empty rows are skipped."
	)]
	public SwitchParameter IncludeEmptyRows { get; set; } = false;

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
					new ArgumentException(
						$"Unsupported file type '{fileExtension}' for '{providerPath}'.", "Path"),
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

				try
				{
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
				catch (ArgumentException ex) when (ex.Message.EndsWith("is not a valid Excel file"))
				{
					Error(
						new InvalidDataException($"{providerPath} has a supported Excel extension but the content is not recognized or unreadable."),
						"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
						"UnknownFileContent",
						providerPath
					);
					continue;
				}
				catch (InvalidOperationException ex) when (ex.Message == "Sequence contains no elements")
				{
					Error(
						new InvalidDataException($"{providerPath} has a supported Excel extension but the content is not recognized or unreadable	(no elements found)."),
						"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
						"UnknownFileContent",
						providerPath
					);
					continue;
				}
				catch (NotSupportedException ex)
				{
					if (!ex.Message.Contains("Stream cannot know the file type"))
					{
						throw;
					}
					Error(
						new InvalidDataException($"{providerPath} has a supported Excel extension but the content is not recognized or unreadable."),
							"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
							"UnknownFileContent",
							providerPath
						);
				}
				catch (Exception ex)
				{
					Error(
						ex,
						"Something went wrong in the underlying MiniExcel library. Please file an issue in the ExcelFast GitHub repository.",
						"MiniExcelError",
						providerPath,
						errorDetailsMessage: $"Error importing '{providerPath}': MiniExcel Query failed: {ex.Message}"
					);
					continue;
				}
			}
			catch (Exception ex)
			{
				Error(
					ex,
					"Something unexpected went wrong while importing the Excel file. Please file an issue in the ExcelFast GitHub repository.",
					"ImportFailed",
					providerPath
				);
				continue;
			}

			if (Raw.IsPresent)
			{
				// Return the raw enumerable as-is so the consumer can stream/transform using their preferred method.
				WriteObject(rows, false);
				continue;
			}

			int rowCount = 0;
			foreach (IDictionary<string, object> row in rows)
			{
				rowCount++;

				if (!IncludeEmptyRows && row.Values.All(v => v == null))
				{
					Debug($"Row {rowCount} in '{providerPath}' sheet '{SheetName}' is empty. Skipping. Specify -IncludeEmptyRows to include null rows.");
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
