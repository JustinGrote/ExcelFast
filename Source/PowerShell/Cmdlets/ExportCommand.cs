namespace ExcelFast.PowerShell.Cmdlets;

using System;
using System.Threading.Channels;

using ExcelFast.Extensions;

using MiniExcelLib;
using MiniExcelLib.OpenXml;

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
  [Alias("Path")]
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

  [Parameter(
    HelpMessage = "Include properties that could not be converted or accessed by exporting a placeholder value describing the failure."
  )]
  public SwitchParameter IncludeUnexportableProperties { get; set; }

  // The running task to export data to Excel
  private Task<int[]>? exportTask;

  private bool whatIfSpecified = false;

  private int rowsExported;

  private readonly Channel<Dictionary<string, object>> exportQueue;

  public ExportCommand()
  {
    exportQueue = Channel.CreateBounded<Dictionary<string, object>>(new BoundedChannelOptions(InputQueueSize));
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

      exportQueue.Writer.TryComplete();
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

      try
      {
        await Post(() =>
        {
          Dictionary<string, string> conversionErrors;
          var row = inputObject.ToFlatDictionary(out conversionErrors, IncludeUnexportableProperties.IsPresent);

          foreach (KeyValuePair<string, string> conversionError in conversionErrors)
          {
            Debug($"Skipping property '{conversionError.Key}' because it {conversionError.Value}");
          }

          exportQueue.Writer.TryWrite(row);
        });

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
      catch (Exception ex)
      {
        Debug($"Unable to convert input object to a row for export: {ex.Message}");
      }
    }
  }

  private void StartExportTaskIfNeeded()
  {
    if (exportTask is not null)
    {
      return;
    }

    var queue = exportQueue.Reader.ReadAllAsync(PipelineStopToken);

    Debug("Starting MiniExcel export task.");
    var exporter = MiniExcel.Exporters.GetOpenXmlExporter();
    exportTask = Task.Run(async () => await exporter.ExportAsync(
      Destination,
      queue,
      sheetName: SheetName,
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