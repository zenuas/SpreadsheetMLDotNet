using SpreadsheetMLDotNet.Data.Styles;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Row : IHaveStyle
{
    public int StartCellIndex { get; set; } = 0;
    public IndexedList<Cell> Cells { get; init; } = new() { New = () => new() { Value = CellValueNull.Instance } };
    public double? Height { get; set; } = null;
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
    public INumberFormat? NumberFormat { get; set; }
}
