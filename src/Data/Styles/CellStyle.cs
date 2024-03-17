namespace SpreadsheetMLDotNet.Data.Styles;

public struct CellStyle
{
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
    public INumberFormat? NumberFormat { get; set; }
}
