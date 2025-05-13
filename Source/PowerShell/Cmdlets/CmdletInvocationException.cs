namespace ExcelFast.PowerShell.Cmdlets;

#pragma warning disable RCS1194 // Implement exception constructors
class CmdletInvocationException(string message) : Exception(message) { }
#pragma warning restore RCS1194 // Implement exception constructors