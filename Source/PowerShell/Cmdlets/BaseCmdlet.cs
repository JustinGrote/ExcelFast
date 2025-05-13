namespace ExcelFast.PowerShell.Cmdlets;

public abstract class BaseCmdlet : PSCmdlet
{
	protected string name => MyInvocation.MyCommand.Name;

	internal void Debug(string message) => WriteDebug($"{name}: {message}");
	internal void Verbose(string message) => WriteVerbose($"{name}: {message}");
	internal void Warning(string message) => WriteWarning($"{name}: {message}");

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
			bool terminating = false) =>
					Error(new CmdletInvocationException(message), recommendedAction, errorId, targetObject, null, category, terminating);
}
