namespace ExcelFast.PowerShell.Cmdlets;

using System.Net.Http;

public abstract class BaseCmdlet : PSCmdlet
{
	private static readonly HttpClient HttpClient = new();

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

	protected string ResolveWorkbookPath(string pathItem, out string? tempPath)
	{
		tempPath = null;

		if (Uri.TryCreate(pathItem, UriKind.Absolute, out Uri? uri) &&
			(uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
		{
			string fileName = Path.GetFileName(uri.LocalPath);
			if (string.IsNullOrWhiteSpace(fileName))
			{
				string extension = Path.GetExtension(uri.AbsolutePath);
				if (string.IsNullOrWhiteSpace(extension))
				{
					extension = ".tmp";
				}

				tempPath = Path.Combine(Path.GetTempPath(), $"ExcelFast-{Guid.NewGuid():N}{extension}");
			}
			else
			{
				tempPath = Path.Combine(Path.GetTempPath(), fileName);
				if (File.Exists(tempPath))
				{
					string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
					string extension = Path.GetExtension(fileName);
					tempPath = Path.Combine(Path.GetTempPath(), $"{fileNameWithoutExtension}-{Guid.NewGuid():N}{extension}");
				}
			}

			Debug($"Downloading workbook from '{uri}' to temporary file '{tempPath}'.");

			using HttpRequestMessage request = new(HttpMethod.Get, uri);
			using HttpResponseMessage response = HttpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
			response.EnsureSuccessStatusCode();

			using Stream contentStream = response.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
			using FileStream outputStream = File.Create(tempPath);
			contentStream.CopyTo(outputStream);

			return tempPath;
		}

		return GetUnresolvedProviderPathFromPSPath(pathItem);
	}
}
