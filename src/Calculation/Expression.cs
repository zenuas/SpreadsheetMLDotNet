namespace SpreadsheetMLDotNet.Calculation;

public class Expression : IFormula
{
    public required string Operator { get; init; }
    public required IFormula Left { get; init; }
    public required IFormula Right { get; init; }
}
