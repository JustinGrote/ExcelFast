namespace ExcelFast.PowerShell.Cmdlets;

using System;
using System.Threading.Channels;
using System.Collections.Concurrent;

using ExcelFast.Extensions;

using MiniExcelLibs;

using FilePath = Path;

[Cmdlet(VerbsData.Export, CmdletDefaultName, SupportsShouldProcess = true)]
[Alias("exwb")]
public class ExportCommand : TaskCmdlet
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

  // The running task to export data to Excel
  private Task<int[]>? exportTask;

  private bool whatIfSpecified = false;

  private int rowsExported;

#if NET472
  private readonly BlockingCollection<Dictionary<string, object>> exportQueue = [];
#else
  private readonly Channel<Dictionary<string, object>> exportQueue
    = Channel.CreateBounded<Dictionary<string, object>>(new BoundedChannelOptions(1));
#endif

  protected override async Task Begin()
  {
    string providerPath = GetUnresolvedProviderPathFromPSPath(Destination);
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
    whatIfSpecified = !await ShouldProcessAsync(providerPath, operation);
    if (whatIfSpecified) return;

    // Create directory if it doesn't exist
    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
    {
      Directory.CreateDirectory(directory);
    }

    // Update destination for processing
    Destination = providerPath;
  }

  protected override async Task Process()
  {
    if (PipelineStopToken.IsCancellationRequested) return;
    if (InputObject is null || InputObject.Length == 0) return;

    await ProcessInputObjects();
  }

  protected override async Task End()
  {
    System.Console.WriteLine("DEBUGTEST: Start End. Export task is " + (exportTask is null ? "null" : exportTask.Status.ToString()));

    if (exportTask is null) await ProcessInputObjects();

    if (whatIfSpecified)
    {
      Info($"WhatIf - Would have written {rowsExported} rows to '{Destination}'", ["PSHOST"]);
      return;
    }

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
      System.Console.WriteLine("DEBUGTEST: Nothing further to queue. Closing Export Queue.");
#if NET472
      exportQueue.CompleteAdding();
#else
      exportQueue.Writer.TryComplete();
#endif
      System.Console.WriteLine("DEBUGTEST: Waiting for export task to complete.");

      Debug("Waiting for export task to complete.");

      int[] result = await exportTask;
      System.Console.WriteLine("DEBUGTEST: Waiting for export task to complete.");

      Verbose($"Exported {result.Sum()} rows to '{Destination}'.");
    }
    catch (OperationCanceledException)
    {
      Warning($"Export cancelled for '{Destination}'.");
    }
    catch (Exception ex)
    {
      System.Console.WriteLine($"DEBUGTEST: Exception in export task: {ex}");
      Error(
        ex,
        "An error occurred during export. See exception details for more information.",
        "ExportFailed",
        Destination,
        terminating: true
      );
    }
  }

  private async Task ProcessInputObjects()
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
        rowsExported++;
        // Dont actually do any work if we are just testing with -WhatIf, but we still want to count the rows for accurate reporting
        if (whatIfSpecified) return;

#if NET472
        exportQueue.Add(row, PipelineStopToken);
#else
        await exportQueue.Writer.WriteAsync(row, PipelineStopToken);
#endif
      }
      catch (OperationCanceledException)
      {
        Debug("Export row enqueue canceled.");
        return;
      }



      // We dont want to start the SaveAsAsync task until we have at least one row to export, to avoid creating an empty file if the input is empty
      if (exportTask is null)
      {
        System.Console.WriteLine("DEBUGTEST: Getting Enumerable");
#if NET472
  var queue = exportQueue.GetConsumingEnumerable(PipelineStopToken);
#else
        var queue = exportQueue.Reader.ReadAllAsync(PipelineStopToken);
#endif
        System.Console.WriteLine("DEBUGTEST: Starting export task.");

        // Start the export task immediately to begin streaming
        Debug("Starting MiniExcel export task.");

        exportTask = Task.Run(async () => await MiniExcel.SaveAsAsync(
          Destination,
          queue,
          sheetName: SheetName,
          excelType: ExcelType.XLSX,
          overwriteFile: Force.IsPresent,
          cancellationToken: PipelineStopToken
        ));
        System.Console.WriteLine("DEBUGTEST: Export task started.");

      }
    }
  }

  protected override async Task Clean()
  {
    System.Console.WriteLine("DEBUG: Clean called, initiating cancellation and cleanup of export task.");
    Debug("Stopping export process due to pipeline stop request.");
#if NET472
    exportQueue.CompleteAdding();
#else
    exportQueue.Writer.TryComplete();
#endif

    if (exportTask is null)
    {
      return;
    }

    try
    {
      Debug("Waiting for export task to acknowledge cancellation.");
      await exportTask;
      Debug($"Export stopped for destination {Destination}.");
    }
		catch (OperationCanceledException)
		{
			Debug("Export task cancellation observed.");
    }
  }

  private async void AssertExportTaskNotFaulted()
  {
    if (exportTask?.IsFaulted == true) await exportTask; // This will re-throw the exception from the export task to be handled by the cmdlet's error handling
  }
}
