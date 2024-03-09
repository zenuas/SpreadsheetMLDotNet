namespace SpreadsheetMLDotNet.Calculation;

public class Unary : IFormula
{
    public required string Operator { get; init; }
    public required IFormula Value { get; init; }
}
