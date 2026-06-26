namespace ExcelFast.PowerShell.Cmdlets;

using ClosedXML.Excel;

[Cmdlet(VerbsCommon.Add, "Worksheet")]
[Alias("aws")]
[OutputType(typeof(IXLWorksheet))]
public class AddWorksheetCommand : BetterPSCmdlet
{
  [Parameter(
    Mandatory = true,
    Position = 0,
    ValueFromPipeline = true,
    HelpMessage = "Workbook where the worksheet should be added."
  )]
  [ValidateNotNull]
  [NotNull]
  public IXLWorkbook? Workbook { get; set; }

  [Parameter(
    Position = 1,
    HelpMessage = "Name of the worksheet to add. If omitted, ClosedXML generates a default sheet name."
  )]
  [ValidateNotNullOrEmpty]
  public string? Name { get; set; }

  [Parameter(
    HelpMessage = "1-based position where the worksheet should be inserted."
  )]
  [ValidateRange(1, int.MaxValue)]
  public int? Position { get; set; }

  protected override void ProcessRecord()
  {
    if (Workbook is null)
    {
      return;
    }

    try
    {
      IXLWorksheet worksheet;
      if (string.IsNullOrWhiteSpace(Name))
      {
        worksheet = Workbook.AddWorksheet();
        if (Position.HasValue)
        {
          worksheet.Position = Position.Value;
        }
      }
      else if (Position.HasValue)
      {
        worksheet = Workbook.AddWorksheet(Name, Position.Value);
      }
      else
      {
        worksheet = Workbook.AddWorksheet(Name);
      }

      WriteObject(worksheet);
    }
    catch (Exception ex)
    {
      Error(
        ex,
        "Verify the worksheet name and position are valid for this workbook.",
        "AddWorksheetFailed",
        Workbook
      );
    }
  }
}
