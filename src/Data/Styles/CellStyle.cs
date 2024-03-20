namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct CellStyle
{
    public Font? Font { get; init; }
    public Fill? Fill { get; init; }
    public Border? Border { get; init; }
    public Alignment? Alignment { get; init; }
    public INumberFormat? NumberFormat { get; init; }
}
