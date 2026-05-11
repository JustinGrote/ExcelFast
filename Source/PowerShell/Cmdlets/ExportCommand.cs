namespace ExcelFast.PowerShell.Cmdlets;

using System.Collections;
using System.Collections.Concurrent;
using System.Threading;

using ExcelFast.Extensions;

using MiniExcelLibs;

using static System.Management.Automation.PSSerializer;

using FilePath = Path;

[Cmdlet(VerbsData.Export, CmdletDefaultName, SupportsShouldProcess = true)]
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
	[NotNull]
	public PSObject[]? InputObject { get; set; }

	[Parameter(
			HelpMessage = "Name of the sheet to export to. If not specified, exports to 'Sheet1'."
	)]
	[ValidateNotNullOrEmpty]
	public string SheetName { get; set; } = "Sheet1";

	[Parameter(
			HelpMessage = "Forces overwriting of the destination file if it already exists."
	)]
	public SwitchParameter Force { get; set; }

	// Collection "queue" to store data to be exported, allowing for concurrent processing
	private readonly BlockingCollection<Dictionary<string, object>> exportQueue = new();
	private readonly CancellationTokenSource exportCancellationTokenSource = new();
	private CancellationToken cancelToken
	{
		get
		{
			exportCancellationTokenSource.Token.ThrowIfCancellationRequested();
			return exportCancellationTokenSource.Token;
		}
	}

	// The running task to export data to Excel
	private Task<int[]>? exportTask;

	private bool whatIfSpecified = false;

	private int rowsExported;

	protected override void BeginProcessing()
	{
		if (Destination is null)
		{
			return;
		}

		string providerPath = GetUnresolvedProviderPathFromPSPath(Destination);
		try
		{
			string fileExtension = FilePath.GetExtension(providerPath).ToLowerInvariant();
			if (!AcceptedExtensions.Contains(fileExtension))
			{
				Error(
					new ArgumentException($"Unsupported file type '{fileExtension}' for '{providerPath}'.", "Path"),
					$"Use one of the supported file types: {string.Join(", ", AcceptedExtensions)}",
					"UnsupportedFileType",
					providerPath,
					terminating: true
				);
				return;
			}

			string? directory = FilePath.GetDirectoryName(providerPath);
			bool directoryExists = string.IsNullOrEmpty(directory) || Directory.Exists(directory);
			bool destinationExists = File.Exists(providerPath);

			// Check if file or directory needs force
			if (!Force.IsPresent && (!directoryExists || destinationExists))
			{
				Error(
					new IOException($"Path '{providerPath}' already exists or requires directory creation."),
					"Use -Force to proceed with the operation.",
					"PathRequiresForce",
					providerPath,
					terminating: true
				);
				return;
			}

			string operation = destinationExists ? "Overwrite Workbook" : "Create Workbook";
			whatIfSpecified = !ShouldProcess(providerPath, operation);
			if (whatIfSpecified) return;

			// Create directory if it doesn't exist
			if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
			{
				Directory.CreateDirectory(directory);
			}

			// Update destination for processing
			Destination = providerPath;
		}
		catch (Exception ex)
		{
			Error(
				ex,
				"This is probably a bug, please report it.",
				"ExportInitializationFailed",
				providerPath,
				terminating: true
			);
		}
	}

	protected override void ProcessRecord()
	{
		AssertExportTaskNotFaulted();
		if (exportCancellationTokenSource.IsCancellationRequested)
		{
			return;
		}

		if (InputObject is null || InputObject.Length == 0) return;

		ProcessInputObjects();
	}

	protected override void EndProcessing()
	{
		if (exportTask is null) ProcessInputObjects();

		if (whatIfSpecified)
		{
			Info($"What if: Would have written {rowsExported} rows to '{Destination}'", ["PSHOST"]);
			return;
		}

		exportQueue.CompleteAdding();

		if (exportTask is null)
		{
			Error(
				new InvalidOperationException("Export task was not initialized."),
				"This is probably a bug, please report it.",
				"ExportTaskNotInitialized",
				terminating: true
			);
			return;
		}

		try
		{
			Debug("Waiting for export task to complete.");
			while (!exportTask.IsCompleted)
			{
				// Enables Ctrl-C to still work while waiting for the export task to complete
				Thread.Sleep(100);
			}
			int[] result = exportTask.GetAwaiter().GetResult();
			Debug($"Exported {result.Sum()} rows to '{Destination}'.");
		}
		catch (OperationCanceledException)
		{
			Warning($"Export cancelled for '{Destination}'.");
		}
		catch (Exception ex)
		{
			Error(
				ex,
				"This is probably a bug, please report it.",
				"ExportProcessFailed",
				terminating: true
			);
		}
	}


	private void ProcessInputObjects()
	{
		foreach (PSObject inputObject in InputObject)
		{
			if (inputObject is null)
			{
				Debug($"Skipping null input object.");
				return;
			}
			Dictionary<string, object> row = inputObject.ToFlatDictionary();
			try
			{
				if (!whatIfSpecified) exportQueue.Add(row, cancelToken);
				rowsExported++;
				if (whatIfSpecified) return;
			}
			catch (OperationCanceledException)
			{
				Debug("Export row enqueue canceled.");
				return;
			}

			// We must start after one item is enqueued, or else SaveAsAsync will hang.
			if (exportTask is null)
			{
				// Start the export task immediately to begin streaming
				Debug("Starting MiniExcel export task.");
				// HACK: SaveAsAsync blocks on the consuming enumerable before returning the task object, so we wrap it in an outer task

				exportTask = StartExporter();
			}
		}
	}

	protected override void StopProcessing()
	{
		Debug("Stopping export process due to pipeline stop request.");
		exportCancellationTokenSource.Cancel();
		exportQueue.CompleteAdding();

		if (exportTask is null)
		{
			return;
		}

		try
		{
			Debug("Waiting for export task to acknowledge cancellation.");
			exportTask.GetAwaiter().GetResult();
			Debug($"Export stopped for destination {Destination}.");
		}
		catch (OperationCanceledException)
		{
			Debug("Export task cancellation observed.");
		}
		catch (Exception ex)
		{
			Error(
				ex,
				"This is probably a bug, please report it.",
				"ExportProcessFailed",
				terminating: true
			);
		}
	}

	private void AssertExportTaskNotFaulted()
	{
		if (exportTask?.IsFaulted == true)
		{
			try
			{
				exportTask.GetAwaiter().GetResult();
			}
			catch (Exception ex)
			{
				Error(
					ex,
					"An error occurred in the export process.",
					"ExportTaskError",
					terminating: true
				);
			}
		}
	}

	private Task<int[]> StartExporter() =>
		Task.Run(async () =>
			await MiniExcel.SaveAsAsync(
				Destination,
				exportQueue.GetConsumingEnumerable(cancelToken),
				sheetName: SheetName,
				excelType: ExcelType.XLSX,
				overwriteFile: Force.IsPresent,
				cancellationToken: cancelToken
			), cancelToken
		);
}
