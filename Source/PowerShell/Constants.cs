namespace ExcelFast;

internal static class Constants
{
	internal const string CmdletPrefix = "";
	internal const string CmdletSuffix = "Workbook";
	internal const string CmdletDefaultName = CmdletPrefix + CmdletSuffix;
	internal static readonly string[] AcceptedExtensions = [".xlsx", ".csv"];
}

internal static class SystemConstants
{
	internal const int ERROR_SHARING_VIOLATION = -2147024864;
}