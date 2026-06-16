namespace System.Management.Automation;

using System.Collections;
using System.Collections.Concurrent;

public abstract class TaskCmdlet : BetterPSCmdlet, IDisposable
{
  // This is our API that your derived class can implement
  protected virtual Task Begin() => Task.CompletedTask;
  protected virtual Task Process() => Task.CompletedTask;
  protected virtual Task End() => Task.CompletedTask;
  protected virtual Task Clean() => Task.CompletedTask;

  /// <summary>
  /// Queues output for the cmdlet pipeline. This is the primary method used to send data through the async pipeline,
  /// and the other output helpers build on top of it.
  /// </summary>
  internal void AddOutput(object item, bool raw = false, CancellationToken? cancelToken = null)
  {
    if (_output == null)
    {
      throw new InvalidOperationException("WriteObject cannot be called before the pipeline is initialized.");
    }
    _output.Add(new(item, raw), cancelToken ?? PipelineStopToken);
  }

  /// <summary>
  /// Buffers output from asynchronous pipeline steps on separate threads before writing it to the pipeline.
  /// </summary>
  /// <remarks>
  /// A fresh collection is created for each pipeline execution, but it must remain accessible to the overridden <c>WriteObject</c> method.
  /// </remarks>
  private BlockingCollection<OutputItem>? _output;

  // Override the pscmdlet entrypoints to execute our async methods
  protected sealed override void BeginProcessing() => ExecuteAsyncPipelineStep(Begin);
  protected sealed override void ProcessRecord() => ExecuteAsyncPipelineStep(Process);
  protected sealed override void EndProcessing() => ExecuteAsyncPipelineStep(End);

  // PowerShell 7.6 introduces a builtin PipelineStopToken, for older versions we
  // implement our own cancellation stop trigger and its cleanup. Once 7.6 is the
  // baseline we can remove the else block.
  private bool _disposed;

  public void Dispose()
  {
    Dispose(true);
    GC.SuppressFinalize(this);
  }

#if NET10_0_OR_GREATER
  protected virtual void Dispose(bool disposing) {}
  protected sealed override void StopProcessing() => ExecuteAsyncPipelineStep(Clean);
#else
  private readonly CancellationTokenSource _cancelSource = new();

  public CancellationToken PipelineStopToken => _cancelSource.Token;

  protected override void StopProcessing()
  {
    _cancelSource.Cancel();
    ExecuteAsyncPipelineStep(Clean);
  }

  protected virtual void Dispose(bool disposing)
  {
    if (_disposed)
    {
      return;
    }

    if (disposing)
    {
      _cancelSource.Dispose();
    }

    _disposed = true;
  }
#endif

  /// <summary>
  /// Executes an asynchronous cmdlet step and routes its output and errors through the PowerShell pipeline.
  /// </summary>
  private void ExecuteAsyncPipelineStep(Func<Task> cmdletMethod)
  {
    _output = [];
    Task.Run(async () =>
    {
      try
      {
        await cmdletMethod();
      }
      catch (Exception ex)
      {
        // Handle exceptions by writing them to the error stream
        Error(ex);
      }
      finally
      {
        _output.CompleteAdding();
      }
    }, PipelineStopToken);

    foreach (var item in _output.GetConsumingEnumerable(PipelineStopToken))
    {
      ProcessOutput(item);
    }
  }

  private void ProcessOutput(OutputItem inputObject)
  {
    var (item, raw) = inputObject;

    if (raw)
    {
      base.WriteObject(item);
      return;
    }

    switch (item)
    {
      case ShouldProcessPrompt prompt:
        bool response = string.IsNullOrEmpty(prompt.Action)
          ? base.ShouldProcess(prompt.Target)
          : base.ShouldProcess(prompt.Target, prompt.Action);
        prompt.Response.TrySetResult(response);
        break;
      case ShouldProcessCustomPrompt customPrompt:
        bool customResponse = ShouldProcess(customPrompt.whatIfMessage, customPrompt.confirmHeader, customPrompt.confirmMessage);
        customPrompt.Response.TrySetResult(customResponse);
        break;
      case ErrorRecord errorRecord:
        base.WriteError(errorRecord);
        break;
      case InformationRecord informationRecord:
        base.WriteInformation(informationRecord);
        break;
      case TaggedInformationInfo taggedInformationInfo:
        base.WriteInformation(taggedInformationInfo.MessageData, taggedInformationInfo.Tags);
        break;
      case WarningRecord warningRecord:
        base.WriteWarning(warningRecord.Message);
        break;
      case VerboseRecord verboseRecord:
        base.WriteVerbose(verboseRecord.Message);
        break;
      case DebugRecord debugRecord:
        base.WriteDebug(debugRecord.Message);
        break;
      case ProgressRecord progressRecord:
        base.WriteProgress(progressRecord);
        break;
      default:
        base.WriteObject(item);
        break;
    }
  }


  // These are compatibility shims for the other Write* methods. You should use the other methods above instead of these, but this allows existing code to work without modification.
  // Override WriteObject to write to our blocking collection instead of the pipeline directly, to avoid marshalling issues
  protected new void WriteObject(object outputObject, bool enumerateCollection = false)
  {
    if (enumerateCollection && outputObject is IEnumerable enumerable)
    {
      foreach (var item in enumerable)
      {
        AddOutput(item, true);
        return;
      }
    }

    AddOutput(outputObject, true);
  }

  protected new void WriteWarning(string message) => AddOutput(new WarningRecord(message));
  protected new void WriteVerbose(string message) => AddOutput(new VerboseRecord(message));
  protected new void WriteDebug(string message) => AddOutput(new DebugRecord(message));
  protected new void WriteError(ErrorRecord errorRecord) => AddOutput(errorRecord);
  protected new void WriteProgress(ProgressRecord progressRecord) => AddOutput(progressRecord);
  protected new void WriteInformation(InformationRecord informationRecord) => AddOutput(informationRecord);
  protected new void WriteInformation(object messageData, string[] tags)
    => AddOutput(new TaggedInformationInfo(messageData, tags));
  protected new bool ShouldProcess(string target, string action = "")
        => ShouldProcessAsync(target, action).GetAwaiter().GetResult();

  protected async Task<bool> ShouldProcessAsync(string target, string action = "")
  {
    TaskCompletionSource<bool> response = new();
    PipelineStopToken.Register(() => response.TrySetCanceled());
    AddOutput(new ShouldProcessPrompt(target, action, response));
    return await response.Task;
  }

  protected async Task<bool> ShouldProcessCustom(string whatIfMessage, string confirmHeader = "", string confirmMessage = "")
  {
    TaskCompletionSource<bool> response = new();
    PipelineStopToken.Register(() => response.TrySetCanceled());
    AddOutput(new ShouldProcessCustomPrompt(whatIfMessage, confirmHeader, confirmMessage, response));
    return await response.Task;
  }
}

public class BetterPSCmdlet : PSCmdlet
{
  protected string name => MyInvocation.MyCommand.Name;

  internal void Debug(string message, bool raw = false) => WriteDebug(raw ? message : $"{name}: {message}");
  internal void Verbose(string message, bool raw = false) => WriteVerbose(raw ? message : $"{name}: {message}");
  internal void Warning(string message, bool raw = false) => WriteWarning(raw ? message : $"{name}: {message}");
  internal void Info(string message, string[]? tags = null, bool raw = false)
    => WriteInformation(raw ? message : $"{name}: {message}", tags ?? []);
  internal void Console(string message, bool raw = false)
    => WriteInformation(raw ? message : $"{name}: {message}", ["PSHOST"]);
  internal void Progress(
    string activity,
    string status = "",
    string currentOperation = "",
    int percentComplete = 0,
    int id = 1,
    int parentId = -1,
    bool completed = false
  ) => WriteProgress(new ProgressRecord(id, activity, status)
  {
    CurrentOperation = currentOperation,
    ParentActivityId = parentId,
    RecordType = completed || percentComplete == 100 ? ProgressRecordType.Completed : ProgressRecordType.Processing,
    PercentComplete = percentComplete
  });

  internal void Error(
      Exception exception,
      string? recommendedAction = null,
      string errorId = "PSCmdletError",
      object? targetObject = null,
      // Usually comes from the exception message, specify this to override
      string? errorDetailsMessage = null,
      // This is often autodetermined
      ErrorCategory? category = null,
      bool terminating = false)
  {
    ErrorRecord error = new(
        exception,
        errorId,
        category ?? exception switch
        {
          ArgumentException => ErrorCategory.InvalidArgument,
          FileNotFoundException => ErrorCategory.ObjectNotFound,
          InvalidOperationException => ErrorCategory.InvalidOperation,
          NotSupportedException => ErrorCategory.NotSpecified,
          UnauthorizedAccessException => ErrorCategory.SecurityError,
          PathTooLongException => ErrorCategory.InvalidArgument,
          DirectoryNotFoundException => ErrorCategory.ObjectNotFound,
          IOException => ErrorCategory.WriteError,
          NullReferenceException => ErrorCategory.InvalidData,
          FormatException => ErrorCategory.InvalidData,
          TimeoutException => ErrorCategory.OperationTimeout,
          OutOfMemoryException => ErrorCategory.ResourceUnavailable,
          NotImplementedException => ErrorCategory.NotImplemented,
          OperationCanceledException => ErrorCategory.OperationStopped,
          AccessViolationException => ErrorCategory.SecurityError,
          InvalidCastException => ErrorCategory.InvalidType,
          _ => ErrorCategory.NotSpecified
        },
        targetObject
    )
    {
      ErrorDetails = new ErrorDetails(errorDetailsMessage ?? exception.Message)
      {
        RecommendedAction = recommendedAction
      }
    };

    if (terminating)
    {
      ThrowTerminatingError(error);
    }
    else
    {
      WriteError(error);
    }
  }

  internal void Error(
    string message,
    string? recommendedAction = null,
    string errorId = "PSCmdletError",
    object? targetObject = null,
    ErrorCategory category = ErrorCategory.NotSpecified,
    bool terminating = false
  ) => Error(
    new CmdletInvocationException(message), recommendedAction, errorId, targetObject, null, category, terminating
  );
}

internal record OutputItem(object Item, bool Raw);
internal record TaggedInformationInfo(object MessageData, string[] Tags);
internal record ShouldProcessPrompt(string Target, string Action, TaskCompletionSource<bool> Response);
internal record ShouldProcessCustomPrompt(
    string whatIfMessage,
    string confirmHeader,
    string confirmMessage,
    TaskCompletionSource<bool> Response
);