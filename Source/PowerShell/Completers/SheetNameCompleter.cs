using System.Collections;
using System.Management.Automation;
using System.Management.Automation.Language;
using MiniExcelLibs;

namespace ExcelFast.PowerShell.Completers;

public class SheetNameCompleter : IArgumentCompleter
{
	public IEnumerable<CompletionResult> CompleteArgument(
			string commandName,
			string parameterName,
			string wordToComplete,
			CommandAst commandAst,
			IDictionary fakeBoundParameters)
	{
		var results = new List<CompletionResult>();

		// Get the Path parameter value
		if (fakeBoundParameters.Contains("Path"))
		{
			string pathValue = string.Empty;

			// Handle different types of path values
			if (fakeBoundParameters["Path"] is string path)
			{
				pathValue = path;
			}
			else if (fakeBoundParameters["Path"] is string[] paths && paths.Length > 0)
			{
				pathValue = paths[0]; // Use the first path for completion
			}
			else if (fakeBoundParameters["Path"] is PSObject psObj)
			{
				pathValue = psObj.ToString();
			}

			if (!string.IsNullOrEmpty(pathValue))
			{
				try
				{
					// Get the PowerShell session and resolve the path
					using var powershell = System.Management.Automation.PowerShell.Create(RunspaceMode.CurrentRunspace);
					var resolvedPaths = powershell.AddCommand("Get-Item")
																			 .AddParameter("Path", pathValue)
																			 .AddParameter("ErrorAction", "SilentlyContinue")
																			 .Invoke();

					if (resolvedPaths.Count > 0 && resolvedPaths[0].Properties["FullName"]?.Value is string fullPath)
					{
						if (File.Exists(fullPath))
						{
							// Get sheet names from the Excel file
							var sheetNames = MiniExcel.GetSheetNames(fullPath);

							// Filter sheet names based on the word to complete (if any)
							foreach (var sheetName in sheetNames)
							{
								if (string.IsNullOrEmpty(wordToComplete) ||
										sheetName.StartsWith(wordToComplete, StringComparison.OrdinalIgnoreCase))
								{
									// Quote the sheet name if it contains spaces
									string completionText = sheetName.Contains(" ")
											? $"\"{sheetName}\""
											: sheetName;

									results.Add(new CompletionResult(
											completionText,
											sheetName,
											CompletionResultType.ParameterValue,
											$"Sheet: {sheetName}"));
								}
							}
						}
					}
				}
				catch (Exception)
				{
					// Silently fail if we can't get sheet names
				}
			}
		}

		return results;
	}
}
