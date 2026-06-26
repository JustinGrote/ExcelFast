namespace ExcelFast.PowerShell.Cmdlets;

using ClosedXML.Excel;

[Cmdlet(VerbsCommon.Move, "Worksheet")]
[Alias("mws")]
[OutputType(typeof(IXLWorksheet))]
public class MoveWorksheetCommand : BetterPSCmdlet
{
  [Parameter(
    Mandatory = true,
    Position = 0,
    ParameterSetName = "ByWorksheet",
    ValueFromPipeline = true,
    HelpMessage = "Worksheet to move."
  )]
  [ValidateNotNull]
  [NotNull]
  public IXLWorksheet? Worksheet { get; set; }

  [Parameter(
    Mandatory = true,
    Position = 0,
    ParameterSetName = "ByName",
    ValueFromPipeline = true,
    HelpMessage = "Workbook containing the worksheet to move."
  )]
  [ValidateNotNull]
  [NotNull]
  public IXLWorkbook? Workbook { get; set; }

  [Parameter(
    Mandatory = true,
    Position = 1,
    ParameterSetName = "ByName",
    HelpMessage = "Name of the worksheet to move."
  )]
  [ValidateNotNullOrEmpty]
  [NotNull]
  public string? Name { get; set; }

  [Parameter(
    Mandatory = true,
    HelpMessage = "1-based destination position for the worksheet."
  )]
  [ValidateRange(1, int.MaxValue)]
  public int Position { get; set; }

  protected override void ProcessRecord()
  {
    try
    {
      IXLWorksheet? worksheet = ParameterSetName switch
      {
        "ByWorksheet" => Worksheet,
        "ByName" when Workbook is not null => Workbook.Worksheets.FirstOrDefault(candidate =>
          string.Equals(candidate.Name, Name, StringComparison.OrdinalIgnoreCase)
        ),
        _ => null
      };

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

      if (Position > worksheet.Workbook.Worksheets.Count)
      {
        Error(
          new ArgumentOutOfRangeException(nameof(Position), $"Worksheet position '{Position}' is out of range."),
          "Specify a position between 1 and the worksheet count.",
          "WorksheetPositionOutOfRange",
          Position
        );
        return;
      }

      worksheet.Position = Position;
      WriteObject(worksheet);
    }
    catch (Exception ex)
    {
      Error(
        ex,
        "Verify the worksheet and destination position are valid for this workbook.",
        "MoveWorksheetFailed",
        Worksheet is not null ? Worksheet : Workbook
      );
    }
  }
}
