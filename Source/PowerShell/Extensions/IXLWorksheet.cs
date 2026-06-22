namespace ExcelFast.Extensions;

using ClosedXML.Excel;

public static class WorksheetExtensions
{
  extension(IXLWorksheet worksheet)
  {
    public bool TryGetTable(string tableName, out IXLTable? table)
    {
      table = null;

      if (string.IsNullOrWhiteSpace(tableName)) return false;

      table = worksheet.Tables.FirstOrDefault(candidate =>
        string.Equals(candidate.Name, tableName, StringComparison.OrdinalIgnoreCase)
      );

      return table is not null;
    }
  }
}
