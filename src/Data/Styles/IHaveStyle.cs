namespace SpreadsheetMLDotNet.Data.Styles;

public interface IHaveStyle
{
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
}
