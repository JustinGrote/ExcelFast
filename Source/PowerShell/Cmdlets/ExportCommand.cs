using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;

using MiniExcelLibs;

using static System.Management.Automation.PSSerializer;
using static ExcelFast.Constants;

using FilePath = System.IO.Path;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Export, CmdletDefaultName)]
[Alias("exwb")]
public class ExportCommand : BaseCmdlet
{
	[Parameter(
			Mandatory = true,
			Position = 0,
			ValueFromPipelineByPropertyName = true,
			HelpMessage = "Path to the Excel file to export to."
	)]
	[ValidateNotNullOrEmpty]
	[NotNull]
	public string? Destination { get; set; }

	[Parameter(
			Mandatory = true,
			Position = 1,
			ValueFromPipeline = true,
			HelpMessage = "Objects to export to the Excel file."
	)]
	[ValidateNotNull]
	public PSObject[]? InputObject { get; set; }

	[Parameter(
			HelpMessage = "Name of the sheet to export to. If not specified, exports to 'Sheet1'."
	)]
	[ValidateNotNullOrWhiteSpace]
	public string SheetName { get; set; } = "Sheet1";

	[Parameter(
			HelpMessage = "Forces overwriting of the destination file if it already exists."
	)]
	public SwitchParameter Force { get; set; }

	// Used in logging
	private string Name => MyInvocation.MyCommand.Name;

	// Collection to store all input objects organized by sheet
	private readonly List<List<PSObject>> sheetObjects = [];
	// Current sheet being processed
	private List<PSObject> currentSheet = [];

	// Used to store the initial detected columns. This is to ensure subsequent columns are not added.
	private List<string>? columns;

	protected override void ProcessRecord()
	{
		if (InputObject is null || InputObject.Length == 0)
		{
			return;
		}

		foreach (PSObject inputObject in InputObject)
		{
			if (inputObject is null)
			{
				WriteDebug($"{Name}: Skipping null input object.");
				continue;
			}

		}
	}

	private void ConvertToDictionary(PSObject inputObject)
	{
		// Check if this is a nested array
		if (inputObject.BaseObject is Array nestedArray)
		{
			// If we already have objects in the current sheet, add it to our sheets collection
			if (currentSheet.Count > 0)
			{
				sheetObjects.Add(currentSheet);
				currentSheet = [];
			}

			// Create a new sheet for this array
			List<PSObject> arraySheet = [];
			foreach (var item in nestedArray)
			{
				if (item is PSObject psObj)
				{
					arraySheet.Add(psObj);
				}
				else
				{
					arraySheet.Add(new PSObject(item));
				}
			}

			if (arraySheet.Count > 0)
			{
				sheetObjects.Add(arraySheet);
			}
		}
		else
		{
			// Regular object, add to current sheet
			currentSheet.Add(inputObject);
		}
	}

	protected override void EndProcessing()
	{
		// Add any remaining objects in currentSheet to sheetObjects
		if (currentSheet.Count > 0)
		{
			sheetObjects.Add(currentSheet);
		}

		// If no sheets have objects, display warning and return
		if (sheetObjects.Count == 0 || sheetObjects.All(sheet => sheet.Count == 0))
		{
			Warning($"No objects to export.");
			return;
		}

		string providerPath = GetUnresolvedProviderPathFromPSPath(Destination);
		Debug($"Exporting to Excel file: {providerPath}");

		try
		{
			string fileExtension = FilePath.GetExtension(providerPath).ToLowerInvariant();
			if (!AcceptedExtensions.Contains(fileExtension))
			{
				Error(
					new ArgumentException($"Unsupported file type '{fileExtension}' for '{providerPath}'.", "Path"),
					$"Use one of the supported file types: {string.Join(", ", AcceptedExtensions)}",
					"UnsupportedFileType",
					providerPath
				);
				return;
			}

			string? directory = FilePath.GetDirectoryName(providerPath);
			bool directoryExists = string.IsNullOrEmpty(directory) || Directory.Exists(directory);

			// Check if file or directory needs force
			if (!Force.IsPresent && (!directoryExists || File.Exists(providerPath)))
			{
				Error(
					new IOException($"Path '{providerPath}' already exists or requires directory creation."),
					"Use -Force to proceed with the operation.",
					"PathRequiresForce",
					providerPath
				);
				return;
			}

			// Create directory if it doesn't exist
			if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
			{
				Directory.CreateDirectory(directory);
			}

			// Prepare data for all sheets
			Dictionary<string, object> sheetsData = new();

			for (int i = 0; i < sheetObjects.Count; i++)
			{
				List<PSObject> sheetData = sheetObjects[i];
				string sheetName = GenerateSheetName(i);

				// Convert PSObjects to a list of dictionaries for this sheet
				List<Dictionary<string, object>> dataToExport = [];
				foreach (PSObject obj in sheetData)
				{
					// Sanitize the PSObject
					PSObject cleanObj = new(Deserialize(Serialize(obj)));

					Dictionary<string, object> row = [];
					foreach (PSPropertyInfo property in cleanObj.Properties)
					{
						row[property.Name] = property.Value ?? string.Empty;
					}
					dataToExport.Add(row);
				}

				sheetsData[sheetName] = dataToExport;
			}

			// Save data to Excel file
			MiniExcel.SaveAs(providerPath, sheetsData, overwriteFile: Force.IsPresent);
			Verbose($"Successfully exported data to '{providerPath}' across {sheetsData.Count} sheets.");
		}
		catch (Exception ex)
		{
			Error(
				ex,
				"Check file permissions and ensure the file is not locked by another process.",
				"ExportFailed",
				providerPath
			);
		}
	}

	// Helper method to generate sheet names based on the base SheetName
	private string GenerateSheetName(int index)
	{
		if (index == 0)
		{
			return SheetName;
		}

		// Check if the SheetName ends with a number
		string baseName = SheetName;
		int startNumber = 1;

		if (int.TryParse(SheetName[^1].ToString(), out int lastDigit))
		{
			// Find how many trailing digits the sheet name has
			int digitCount = 0;
			for (int i = SheetName.Length - 1; i >= 0; i--)
			{
				if (char.IsDigit(SheetName[i]))
				{
					digitCount++;
				}
				else
				{
					break;
				}
			}

			if (digitCount > 0)
			{
				string numberPart = SheetName.Substring(SheetName.Length - digitCount);
				if (int.TryParse(numberPart, out int number))
				{
					baseName = SheetName.Substring(0, SheetName.Length - digitCount);
					startNumber = number;
				}
			}
		}

		return $"{baseName}{startNumber + index}";
	}
}
