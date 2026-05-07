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
		Position = 0,
		Mandatory = true,
		ParameterSetName = nameof(Workbook),
		ValueFromPipeline = true,
		HelpMessage = "Workbook object to import. Get using Get-Workbook."
	)]
	[ValidateNotNullOrEmpty]
	public IXLWorkbook? Workbook { get; set; }

	[Parameter(
		Position = 0,
		Mandatory = true,
		ParameterSetName = nameof(Range),
		ValueFromPipeline = true,
		HelpMessage = "Range to import. Accepts Table Ranges, Worksheet Ranges, or Workbook Ranges. Get using Get-Workbook, select the appropriate Worksheet, and then select the appropriate Range from the Ranges property."
	)]
	[ValidateNotNullOrEmpty]
	public IXLRangeBase? Range { get; set; }

	[Parameter(
	ParameterSetName = nameof(Path),
	Mandatory = true,
	Position = 0,
	ValueFromPipeline = true,
	ValueFromPipelineByPropertyName = true,
	HelpMessage = "Path to the Excel file to import."
	)]
	[NotNull]
	public string[]? Path { get; set; }

	[Parameter(
		Position = 1,
		HelpMessage = "Names of sheet(s) to import. If not specified, imports the first sheet."
	)]
	public string[]? SheetName { get; set; }

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
			case nameof(Range):
				Path = [GetWorkbookPath(Range!.Worksheet.Workbook)];
				SheetName = [Range.Worksheet.Name];
				StartCell = Range.RangeAddress.FirstAddress.ToStringRelative();
				EndCell = Range.RangeAddress.LastAddress.ToStringRelative();
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
			Error(
				new FileNotFoundException($"Excel file not found: {providerPath}"),
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

		IEnumerable<string> sheetNamesToImport = [];
		try
		{
			IReadOnlyCollection<string> availableSheetNames = MiniExcel.GetSheetNames(providerPath).ToArray();
			if (SheetName == null || SheetName.Length == 0)
			{
				Debug($"No sheet name provided. Importing the first sheet from '{providerPath}'.");
				sheetNamesToImport = [availableSheetNames.FirstOrDefault()];
			}
			else
			{
				IReadOnlyCollection<string> missingSheetNames =
				[
					.. SheetName.Where(sheetName => !availableSheetNames.Contains(sheetName, StringComparer.OrdinalIgnoreCase))
				];

				if (missingSheetNames.Count > 0)
				{
					string missingSheets = string.Join(", ", missingSheetNames);
					Error(
						new ArgumentException($"Sheet(s) '{missingSheets}' do not exist in the '{providerPath}' workbook."),
						"Check the sheet name and try again.",
						"InvalidSheetName",
						string.Join(",", SheetName)
					);
					return;
				}

				sheetNamesToImport = SheetName;
			}

			foreach (string currentSheetName in sheetNamesToImport)
			{
				ImportSheetData(providerPath, currentSheetName);
			}
		}
		catch (ArgumentException ex) when (ex.Message.EndsWith("is not a valid Excel file"))
		{
			Error(
				new InvalidDataException($"{providerPath} has a supported Excel extension but the content is not recognized or unreadable."),
				"The file does not appear to be an Excel file type. Check the extension and content of the file. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
				"UnsupportedFileType",
				providerPath
			);
		}
		catch (InvalidDataException ex) when (ex.Message == "End of Central Directory record could not be found.")
		{
			Error(
				new InvalidDataException($"{providerPath} appears to be a supported Excel file extension but is incomplete or corrupted."),
				"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
				"CorruptedZipContent",
				providerPath
			);
		}
		catch (InvalidOperationException ex) when (ex.Message == "Sequence contains no matching element")
		{
			Error(
				new InvalidDataException($"{providerPath} has a supported Excel extension but the content is not recognized or unreadable	(no elements found)."),
				"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
				"UnknownFileContent",
				providerPath
			);
		}
		catch (InvalidDataException ex) when (ex.Message.StartsWith("The file type could not be inferred automatically"))
		{
			Error(
				new InvalidDataException(ex.Message),
				"The file may be corrupted or not a supported Excel content type. Try opening the file in Excel. If it works, please file an issue in the ExcelFast GitHub repository.",
				"UnknownFileContent",
				providerPath
			);
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
				"Something unexpected went wrong while importing the Excel file. Please file an issue in the ExcelFast GitHub repository.",
				"ImportFailed",
				providerPath,
				terminating: true
			);
		}
	}

	private void ImportSheetData(string providerPath, string currentSheetName)
	{
		IEnumerable<dynamic> rows;

		ICollection<string> columns = MiniExcel.GetColumns(
			providerPath,
			!NoHeaders.IsPresent,
			currentSheetName,
			startCell: StartCell
		);

		bool duplicateColumnSet = !columnSets.Add(columns);

		if (duplicateColumnSet || !columnSets.Any(c => c.SequenceEqual(columns)))
		{
			Warning($"Sheet '{currentSheetName}' in '{providerPath}' has different columns than previously imported sheets. The resultant object output may be different and not displayed correctly.");
		}

		rows = string.IsNullOrEmpty(EndCell)
		? MiniExcel.Query(
			providerPath,
			useHeaderRow: !NoHeaders.IsPresent,
			sheetName: currentSheetName,
			startCell: StartCell
		)
		: MiniExcel.QueryRange(
			providerPath,
			!NoHeaders.IsPresent,
			currentSheetName,
			startCell: StartCell,
			endCell: EndCell
		);

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
				string sheetDisplayName = currentSheetName ?? "<first sheet>";
				Debug($"Row {rowCount} in '{providerPath}' sheet '{sheetDisplayName}' is empty. Skipping. Specify -IncludeEmptyRows to include null rows.");
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

	static string GetWorkbookPath(IXLWorkbook workbook)
	{
		string path = workbook.ToString();
		path = Regex.Replace(path, @"^XLWorkbook\((.*)\)$", "$1");
		return path;
	}
}