namespace ExcelFast.PowerShell.Cmdlets;

using System.Threading.Channels;

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

  private readonly Channel<Dictionary<string, object>> exportQueue
  = Channel.CreateBounded<Dictionary<string, object>>(new BoundedChannelOptions(1));

  protected override async Task Begin()
  {

    string coreLibPath = typeof(object).Assembly.Location;

    if (string.IsNullOrEmpty(coreLibPath) || !System.IO.Path.IsPathFullyQualified(coreLibPath))
    {
      System.Console.WriteLine("CoreLib is not in a rooted path ('{0}')", coreLibPath);
    }
    else
    {
      string? dotnetRuntimeDirectory = System.IO.Path.GetDirectoryName(coreLibPath);
      if (dotnetRuntimeDirectory is null)
      {
        System.Console.WriteLine(".NET Runtime directory is null");
      }
      else
      {
        string? nativeLibraryPrefix = null, nativeLibraryExtension = null;
        if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
        {
          nativeLibraryPrefix = string.Empty;
          nativeLibraryExtension = ".dll";
        }
        else if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux))
        {
          nativeLibraryPrefix = "lib";
          nativeLibraryExtension = ".so";
        }
        else if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX))
        {
          nativeLibraryPrefix = "lib";
          nativeLibraryExtension = ".so";
        }
        else
        {
          System.Console.WriteLine("Unsupported OS");
        }

        if (nativeLibraryPrefix is not null)
        {
          string dbiPath = System.IO.Path.Combine(dotnetRuntimeDirectory, nativeLibraryPrefix + "mscordbi" + nativeLibraryExtension);
          string dacPath = System.IO.Path.Combine(dotnetRuntimeDirectory, nativeLibraryPrefix + "mscordaccore" + nativeLibraryExtension);
          if (!System.IO.File.Exists(dbiPath))
          {
            System.Console.WriteLine("DBI not found at '{0}'", dbiPath);
          }
          else if (!System.IO.File.Exists(dacPath))
          {
            System.Console.WriteLine("DAC not found at '{0}'", dacPath);
          }
          else
          {
            System.Console.WriteLine(".NET Debugging Services libries were found");
          }
        }
      }
    }

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

  protected override async Task Process()
  {
    AssertExportTaskNotFaulted();
    if (PipelineStopToken.IsCancellationRequested) return;
    if (InputObject is null || InputObject.Length == 0) return;

    await ProcessInputObjects();
  }

  protected override async Task End()
  {
    if (exportTask is null) await ProcessInputObjects();

    if (whatIfSpecified)
    {
      Info($"What if: Would have written {rowsExported} rows to '{Destination}'", ["PSHOST"]);
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
      if (!exportQueue.Writer.TryComplete()) Warning("Export queue was already marked as complete, probably a bug.");
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
				"This is probably a bug, please report it.",
				"ExportProcessFailed",
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
        await exportQueue.Writer.WriteAsync(row, PipelineStopToken);
        rowsExported++;
      }
      catch (OperationCanceledException)
      {
        Debug("Export row enqueue canceled.");
        return;
      }

      // We dont want to start the SaveAsAsync task until we have at least one row to export, to avoid creating an empty file if the input is empty
      if (exportTask is null && !whatIfSpecified)
      {
        // Start the export task immediately to begin streaming
        Debug("Starting MiniExcel export task.");

        exportTask = MiniExcel.SaveAsAsync(
          Destination,
          exportQueue.Reader.ReadAllAsync(PipelineStopToken),
          sheetName: SheetName,
          excelType: ExcelType.XLSX,
          overwriteFile: Force.IsPresent,
          cancellationToken: PipelineStopToken
        );
      }
    }
  }

  protected override async Task Clean()
  {
    Debug("Stopping export process due to pipeline stop request.");
    exportQueue.Writer.TryComplete();

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
}
