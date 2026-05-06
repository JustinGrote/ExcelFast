namespace ExcelFast.PowerShell.Cmdlets;

using System.Text.RegularExpressions;

using ClosedXML.Excel;

using DocumentFormat.OpenXml.Drawing.Diagrams;

using MiniExcelLibs;

using FilePath = Path;

[Cmdlet(VerbsData.Import, CmdletDefaultName)]
[Alias("iwb")]
public class ImportCommand : BaseCmdlet
{
	[Parameter(
		ParameterSetName = nameof(Path),
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
		ParameterSetName = nameof(Workbook),
		ValueFromPipeline = true,
		HelpMessage = "Names of sheets to import. If not specified, imports the first sheet."
	)]
	[ValidateNotNullOrEmpty]
	public ClosedXML.Excel.XLWorkbook? Workbook { get; set; }

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
		HelpMessage = "Return the result as a raw dynamic enumerable without PSObject wrapping. Use only for advanced performance use cases."
	)]
	public SwitchParameter Raw { get; set; } = false;

	[Parameter(
		HelpMessage = "Include empty rows in the output. By default, empty rows are skipped."
	)]
	public SwitchParameter IncludeEmptyRows { get; set; } = false;

	readonly HashSet<ICollection<string>> columnSets = [];
	private const string ImportedPSTypeName = "ExcelFast.ImportedWorkbook";

	protected override void ProcessRecord()
	{
		switch (ParameterSetName)
		{
			case nameof(Path):
				// Path is already set from parameter binding, no action needed.
				break;
			case nameof(Workbook):
				Path = [GetWorkbookPath(Workbook!)];
				break;
			default:
				Error(
					new InvalidOperationException("Invalid parameter set: " + ParameterSetName),
					"An unexpected error occurred. Please file an issue in the ExcelFast GitHub repository.",
					"InvalidParameterSet"
				);
				break;
		}
		foreach (string pathItem in Path)
		{
			ImportWorkbook(pathItem);
		}
	}

	internal void ImportWorkbook(string workbookPath)
	{
		string providerPath = GetUnresolvedProviderPathFromPSPath(workbookPath);
		Debug($"Importing Workbook: {providerPath}");

		if (!File.Exists(providerPath))
		{
			Error(new FileNotFoundException($"Excel file not found: {providerPath}"),
				"Check the file path and try again.",
				"FileNotFound",
				providerPath
			);
			return;
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
			return;
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
				return;
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
				return;
			}
			catch (InvalidOperationException ex) when (ex.Message == "Sequence contains no elements")
			{
				Error(
					new InvalidDataException($"{providerPath} has a supported Excel extension but the content is not recognized or unreadable	(no elements found)."),
					"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
					"UnknownFileContent",
					providerPath
				);
				return;
			}
			catch (InvalidDataException ex) when (ex.Message.StartsWith("The file type could not be inferred automatically"))
			{
				Error(
					new InvalidDataException(ex.Message),
					"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
					"UnknownFileContent",
					providerPath
				);
				return;
			}
			catch (NotSupportedException ex) when (ex.Message.Contains("Stream cannot know the file type"))
			{
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
				return;
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
			return;
		}

		if (Raw.IsPresent)
		{
			// Return the raw enumerable as-is so the consumer can stream/transform using their preferred method.
			WriteObject(rows, false);
			return;
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

	static string GetWorkbookPath(XLWorkbook workbook)
	{
		string path = workbook.ToString();
		path = Regex.Replace(path, @"^XLWorkbook\((.*)\)$", "$1");
		return path;
	}
}
