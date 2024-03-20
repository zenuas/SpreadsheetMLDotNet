using SpreadsheetMLDotNet.Data.Styles;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Column : IHaveStyle
{
    public double? Width { get; set; }
    public bool? Hidden { get; set; }
    public bool? BestFitColumnWidth { get; set; }
    public bool? ShowPhonetic { get; set; }
    public uint? OutlineLevel { get; set; }
    public bool? Collapsed { get; set; }
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
    public INumberFormat? NumberFormat { get; set; }
}
