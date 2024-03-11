namespace SpreadsheetMLDotNet.Calculation;

public readonly struct Expression : IFormula
{
    public required string Operator { get; init; }
    public required IFormula Left { get; init; }
    public required IFormula Right { get; init; }
}
