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

  [Parameter(
    HelpMessage = "Bounded queue capacity used to stream rows to the Excel writer. Larger values reduce producer stalls for large exports."
  )]
  public int InputQueueSize { get; set; } = 1024;

  // The running task to export data to Excel
  private Task<int[]>? exportTask;

  private bool whatIfSpecified = false;

  private int rowsExported;

  // #if NET472
  //   private readonly BlockingCollection<Dictionary<string, object>> exportQueue;
  // #else
  private readonly Channel<Dictionary<string, object>> exportQueue;
  // #endif

  public ExportCommand()
  {
    // #if NET472
    //     exportQueue = [];
    // #else
    exportQueue = Channel.CreateBounded<Dictionary<string, object>>(new BoundedChannelOptions(InputQueueSize));
    // #endif
  }

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

    if (exportTask is null)
    {
      await ProcessInputObjects();
    }

    if (whatIfSpecified)
    {
      Info($"WhatIf - Would have written {rowsExported} rows to '{Destination}'", ["PSHOST"]);
      return;
    }

    if (exportTask is null)
    {
      Verbose($"No rows were exported to '{Destination}'.");
      return;
    }

    try
    {
      // #if NET472
      //       exportQueue.CompleteAdding();
      // #else
      exportQueue.Writer.TryComplete();
      // #endif

      Debug("Waiting for export task to complete.");

      int[] result = await exportTask;

      Verbose($"Exported {result.Sum()} rows to '{Destination}'.");
    }
    catch (OperationCanceledException)
    {
      Warning($"Export cancelled for '{Destination}'.");
    }
    catch (Exception ex)
    {
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
    if (InputObject is null || InputObject.Length == 0)
    {
      Debug("No rows were supplied for export.");
      return;
    }

    foreach (PSObject inputObject in InputObject)
    {
      if (inputObject is null)
      {
        Debug("Skipping null input object.");
        continue;
      }

      Exec(() =>
      {
        var row = inputObject.ToFlatDictionary();
        // #if NET472
        //         exportQueue.Add(row, PipelineStopToken);
        // #else
        exportQueue.Writer.TryWrite(row);
        // #endif
      });

      try
      {
        rowsExported++;
        if (whatIfSpecified)
        {
          continue;
        }

        StartExportTaskIfNeeded();
      }
      catch (OperationCanceledException)
      {
        Debug("Export row enqueue canceled.");
        return;
      }
    }
  }

  private void StartExportTaskIfNeeded()
  {
    if (exportTask is not null)
    {
      return;
    }

    // #if NET472
    //     var queue = exportQueue.GetConsumingEnumerable(PipelineStopToken);
    // #else
    var queue = exportQueue.Reader.ReadAllAsync(PipelineStopToken);
    // #endif

    Debug("Starting MiniExcel export task.");

    exportTask = Task.Run(async () => await MiniExcel.SaveAsAsync(
      Destination,
      queue,
      sheetName: SheetName,
      excelType: ExcelType.XLSX,
      overwriteFile: Force.IsPresent,
      cancellationToken: PipelineStopToken
    ), PipelineStopToken);
  }

  protected override async Task Clean()
  {
    Debug("Stopping export process due to pipeline stop request.");
    // #if NET472
    //     exportQueue.CompleteAdding();
    // #else
    exportQueue.Writer.TryComplete();
    // #endif

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