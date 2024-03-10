namespace SpreadsheetMLDotNet.Calculation;

public class Expression : IFormula
{
    public required string Operator { get; set; }
    public required IFormula Left { get; set; }
    public required IFormula Right { get; set; }
}
