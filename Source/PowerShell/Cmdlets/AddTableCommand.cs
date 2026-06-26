namespace ExcelFast.PowerShell.Cmdlets;

using ClosedXML.Excel;

[Cmdlet(VerbsCommon.Add, "Table")]
[Alias("atbl")]
[OutputType(typeof(IXLTable))]
public class AddTableCommand : BetterPSCmdlet
{
  [Parameter(
    Mandatory = true,
    Position = 0,
    ParameterSetName = "FromWorksheet",
    ValueFromPipeline = true,
    HelpMessage = "Worksheet where a table should be created."
  )]
  [ValidateNotNull]
  [NotNull]
  public IXLWorksheet? Worksheet { get; set; }

  [Parameter(
    Mandatory = true,
    Position = 1,
    ParameterSetName = "FromWorksheet",
    HelpMessage = "A1-style range address (e.g. 'A1:B6') to convert to a table."
  )]
  [ValidateNotNullOrEmpty]
  [NotNull]
  public string? RangeAddress { get; set; }

  [Parameter(
    Mandatory = true,
    Position = 0,
    ParameterSetName = "FromRange",
    ValueFromPipeline = true,
    HelpMessage = "Range to convert to a table."
  )]
  [ValidateNotNull]
  [NotNull]
  public IXLRange? Range { get; set; }

  [Parameter(
    HelpMessage = "Name of the created table."
  )]
  [ValidateNotNullOrEmpty]
  public string? Name { get; set; }

  [Parameter(
    HelpMessage = "ClosedXML table theme to apply to the new table."
  )]
  [ValidateNotNullOrEmpty]
  [ArgumentCompleter(typeof(TableStyleArgumentCompleter))]
  public string? Theme { get; set; }

  [Parameter(HelpMessage = "Show totals row on the table.")]
  public SwitchParameter ShowTotalsRow { get; set; }

  [Parameter(HelpMessage = "Hide row stripes on the table.")]
  public SwitchParameter HideRowStripes { get; set; }

  [Parameter(HelpMessage = "Show column stripes on the table.")]
  public SwitchParameter ShowColumnStripes { get; set; }

  [Parameter(HelpMessage = "Hide the table header row.")]
  public SwitchParameter HideHeaderRow { get; set; }

  protected override void ProcessRecord()
  {
    try
    {
      IXLRange tableRange = ParameterSetName switch
      {
        "FromWorksheet" => Worksheet!.Range(RangeAddress!),
        "FromRange" => Range!,
        _ => throw new InvalidOperationException($"Unexpected parameter set '{ParameterSetName}'.")
      };

      IXLTable table = string.IsNullOrWhiteSpace(Name)
        ? tableRange.CreateTable()
        : tableRange.CreateTable(Name);

      if (!string.IsNullOrWhiteSpace(Theme))
      {
        XLTableTheme? parsedTheme = XLTableTheme.GetAllThemes().FirstOrDefault(currentTheme =>
          string.Equals(currentTheme.Name, Theme, StringComparison.OrdinalIgnoreCase)
        );

        if (parsedTheme is null)
        {
          Error(
            new ArgumentException($"The table theme '{Theme}' is not supported by ClosedXML.", nameof(Theme)),
            "Use tab completion to discover supported ClosedXML table themes.",
            "InvalidTableTheme",
            Theme
          );
          return;
        }

        table.Theme = parsedTheme;
      }

      if (HideHeaderRow.IsPresent)
      {
        table.SetShowHeaderRow(false);
      }

      if (ShowTotalsRow.IsPresent)
      {
        table.SetShowTotalsRow(true);
      }

      if (HideRowStripes.IsPresent)
      {
        table.SetShowRowStripes(false);
      }

      if (ShowColumnStripes.IsPresent)
      {
        table.SetShowColumnStripes(true);
      }

      WriteObject(table);
    }
    catch (Exception ex)
    {
      Error(
        ex,
        "Verify the worksheet range and table options are valid.",
        "AddTableFailed",
        Worksheet is not null ? Worksheet : Range
      );
    }
  }
}
