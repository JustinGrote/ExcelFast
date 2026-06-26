namespace ExcelFast.PowerShell.Cmdlets;

using ClosedXML.Excel;

[Cmdlet(VerbsCommon.Remove, "Worksheet", SupportsShouldProcess = true)]
[Alias("rws")]
[OutputType(typeof(IXLWorkbook))]
public class RemoveWorksheetCommand : BetterPSCmdlet
{
  [Parameter(
    Mandatory = true,
    Position = 0,
    ValueFromPipeline = true,
    HelpMessage = "Workbook where the worksheet should be removed."
  )]
  [ValidateNotNull]
  [NotNull]
  public IXLWorkbook? Workbook { get; set; }

  [Parameter(
    Mandatory = true,
    Position = 1,
    ParameterSetName = "ByName",
    HelpMessage = "Name of the worksheet to remove."
  )]
  [ValidateNotNullOrEmpty]
  [NotNull]
  public string? Name { get; set; }

  [Parameter(
    Mandatory = true,
    Position = 1,
    ParameterSetName = "ByPosition",
    HelpMessage = "1-based position of the worksheet to remove."
  )]
  [ValidateRange(1, int.MaxValue)]
  public int Position { get; set; }

  protected override void ProcessRecord()
  {
    if (Workbook is null)
    {
      return;
    }

    try
    {
      switch (ParameterSetName)
      {
        case "ByName":
          IXLWorksheet? worksheet = Workbook.Worksheets.FirstOrDefault(candidate =>
            string.Equals(candidate.Name, Name, StringComparison.OrdinalIgnoreCase)
          );
          if (worksheet is null)
          {
            Error(
              new ArgumentException($"Worksheet '{Name}' was not found in the workbook."),
              "Use Get-Workbook and inspect worksheet names, then retry with an existing worksheet.",
              "WorksheetNotFound",
              Name
            );
            return;
          }

          if (ShouldProcess(worksheet.Name, "Remove Worksheet"))
          {
            Workbook.Worksheets.Delete(worksheet.Name);
          }
          break;
        case "ByPosition":
          if (Position > Workbook.Worksheets.Count)
          {
            Error(
              new ArgumentOutOfRangeException(nameof(Position), $"Worksheet position '{Position}' is out of range."),
              "Specify a position between 1 and the worksheet count.",
              "WorksheetPositionOutOfRange",
              Position
            );
            return;
          }

          if (ShouldProcess($"Position {Position}", "Remove Worksheet"))
          {
            Workbook.Worksheets.Delete(Position);
          }
          break;
        default:
          Error(
            new InvalidOperationException($"Unexpected parameter set '{ParameterSetName}'."),
            "Please file an issue in the ExcelFast repository.",
            "InvalidParameterSet"
          );
          return;
      }

      WriteObject(Workbook);
    }
    catch (Exception ex)
    {
      Error(
        ex,
        "Verify the worksheet name or position is valid for this workbook.",
        "RemoveWorksheetFailed",
        Workbook
      );
    }
  }
}
