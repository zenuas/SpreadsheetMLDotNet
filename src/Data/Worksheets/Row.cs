using SpreadsheetMLDotNet.Data.Styles;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Row : IHaveStyle
{
    public IndexedList<Cell> Cells { get; init; } = new() { New = () => new() { Value = CellValueNull.Instance } };
    public double? Height { get; set; } = null;
    public bool? Hidden { get; set; }
    public uint? OutlineLevel { get; set; }
    public bool? Collapsed { get; set; }
    public bool? ThickTop { get; set; }
    public bool? ThickBottom { get; set; }
    public bool? ShowPhonetic { get; set; }
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
    public INumberFormat? NumberFormat { get; set; }
}
