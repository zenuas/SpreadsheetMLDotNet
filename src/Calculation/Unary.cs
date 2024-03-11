namespace SpreadsheetMLDotNet.Calculation;

public readonly struct Unary : IFormula
{
    public required string Operator { get; init; }
    public required IFormula Value { get; init; }
}
