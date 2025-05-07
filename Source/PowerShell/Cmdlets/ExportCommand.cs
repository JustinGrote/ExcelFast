using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;

using MiniExcelLibs;

using static ExcelFast.Constants;

using FilePath = System.IO.Path;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Export, CmdletDefaultName)]
public class ExportCommand : PSCmdlet
{
	[Parameter(
			Mandatory = true,
			Position = 0,
			ValueFromPipelineByPropertyName = true,
			HelpMessage = "Path to the Excel file to export to."
	)]
	[ValidateNotNullOrWhiteSpace]
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

	// Used in logging
	private string Name => MyInvocation.MyCommand.Name;

	// Collection to store all input objects before writing to Excel
	private readonly List<PSObject> inputObjects = [];

	protected override void ProcessRecord()
	{
		if (InputObject == null || InputObject.Length == 0)
		{
			return;
		}

		foreach (PSObject inputObject in InputObject)
		{
			if (inputObject != null)
			{
				inputObjects.Add(inputObject);
			}
		}
	}

	protected override void EndProcessing()
	{
		if (inputObjects.Count == 0)
		{
			WriteWarning($"No objects to export.");
			return;
		}

		string providerPath = GetUnresolvedProviderPathFromPSPath(Destination);
		WriteDebug($"{Name}: Exporting to Excel file: {providerPath}");

		// Convert PSObjects to a list of dictionaries
		List<Dictionary<string, object>> dataToExport = [];
		foreach (PSObject obj in inputObjects)
		{
			Dictionary<string, object> row = [];
			foreach (PSPropertyInfo property in obj.Properties)
			{
				row[property.Name] = property.Value ?? string.Empty;
			}
			dataToExport.Add(row);
		}

		try
		{
			string fileExtension = FilePath.GetExtension(providerPath).ToLowerInvariant();
			if (!AcceptedExtensions.Contains(fileExtension))
			{
				WriteError(new ErrorRecord(
						new ArgumentException($"Unsupported file type '{fileExtension}' for '{providerPath}'. Supported file types: {string.Join(',', AcceptedExtensions)}", "Path"),
						"UnsupportedFileType",
						ErrorCategory.InvalidArgument,
						providerPath
				));
				return;
			}

			// Create directory if it doesn't exist
			string? directory = FilePath.GetDirectoryName(providerPath);
			if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
			{
				Directory.CreateDirectory(directory);
			}

			// Save data to Excel file
			Dictionary<string, object> value = new()
			{
				[SheetName] = dataToExport
			};

			MiniExcel.SaveAs(providerPath, value);
			WriteVerbose($"Successfully exported {dataToExport.Count} objects to '{providerPath}' in sheet '{SheetName}'.");
		}
		catch (Exception ex)
		{
			WriteError(new ErrorRecord(
					ex,
					"ExportFailed",
					ErrorCategory.WriteError,
					providerPath
			));
		}
	}
}
