using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;

using ClosedXML.Excel;

using static ExcelFast.Constants;

namespace ExcelFast.PowerShell.Cmdlets;

[Cmdlet(VerbsData.Save, CmdletDefaultName)]
[Alias("svwb")]
public class SaveCommand : PSCmdlet
{
	[Parameter(
					Mandatory = true,
					Position = 0,
					ValueFromPipeline = true,
					HelpMessage = "The workbook to save."
	)]
	[NotNull]
	public XLWorkbook? Workbook { get; set; }

	[Parameter(
					Position = 1,
					HelpMessage = "Destination where the Excel file will be saved. If not specified, the workbook will be saved to its current location."
	)]
	[ValidateNotNullOrEmpty]
	[NotNull]
	public string? Destination { get; set; }

	[Parameter(
					HelpMessage = "If specified, overwrites the file if it exists."
	)]
	public SwitchParameter Force { get; set; }

	// Used in logging
	string name => MyInvocation.MyCommand.Name;

	// List to collect workbooks from pipeline
	private readonly List<XLWorkbook> _workbooks = [];

	protected override void ProcessRecord()
	{
		if (Workbook is not null)
		{
			_workbooks.Add(Workbook);
		}
	}

	protected override void EndProcessing()
	{
		// Validate multiple workbooks with Destination scenario
		if (_workbooks.Count > 1 && !string.IsNullOrEmpty(Destination))
		{
			ErrorRecord error = new(
					new PSNotSupportedException(
						"The Destination parameter can only be used when saving a single workbook. " +
						"When saving multiple workbooks, each must be saved to its current location."),
						"MultipleWorkbooksWithDestinationParameter",
					ErrorCategory.InvalidOperation,
					Destination
			);
			error.ErrorDetails.RecommendedAction = "Use -Destination with a single workbook or use a loop to specify the destination separately.";

			WriteError(new ErrorRecord(
					new PSNotSupportedException(
						"The Destination parameter can only be used when saving a single workbook. " +
						"When saving multiple workbooks, each must be saved to its current location."),
						"MultipleWorkbooksWithDestinationParameter",
					ErrorCategory.InvalidOperation,
					Destination
			));
			return;
		}

		// Process each collected workbook
		foreach (XLWorkbook workbook in _workbooks)
		{
			try
			{
				// If no destination is provided, use Save to the current location
				if (string.IsNullOrEmpty(Destination))
				{
					workbook.Save();
					WriteVerbose("Workbook saved to its current location");
					continue;
				}

				// Otherwise, save to the specified destination
				string resolvedPath = GetUnresolvedProviderPathFromPSPath(Destination);

				if (File.Exists(resolvedPath) && !Force.IsPresent)
				{
					WriteError(new ErrorRecord(
							new IOException($"File already exists: {resolvedPath}. Use -Force to overwrite."),
							"FileAlreadyExists",
							ErrorCategory.ResourceExists,
							resolvedPath
					));
					return;
				}

				// Ensure directory exists
				string? directory = Path.GetDirectoryName(resolvedPath);
				if (directory != null && !Directory.Exists(directory))
				{
					if (!Force.IsPresent)
					{
						ErrorRecord error = new(
								new DirectoryNotFoundException($"Directory does not exist: {directory}"),
								"DirectoryNotFound",
								ErrorCategory.InvalidArgument,
								directory
						);
						error.ErrorDetails.RecommendedAction = "Use -Force to create the directory path.";
						WriteError(error);
						return;
					}

					Directory.CreateDirectory(directory);
				}

				workbook.SaveAs(resolvedPath);
				WriteVerbose($"Workbook saved to: {resolvedPath}");
			}
			catch (Exception ex)
			{
				WriteError(new ErrorRecord(
						ex,
						"SaveExcelWorkbookError",
						ErrorCategory.WriteError,
						Destination ?? "current location"
				));
			}
		}
	}
}
