namespace ExcelFast.PowerShell.Cmdlets;

#pragma warning disable RCS1194 // Implement exception constructors
class CmdletInvocationException(string message, Exception? innerException = null) : Exception(message, innerException) { }
#pragma warning restore RCS1194 // Implement exception constructors