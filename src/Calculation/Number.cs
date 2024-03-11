namespace SpreadsheetMLDotNet.Calculation;

public readonly struct Number : IFormula
{
    public required double Value { get; init; }
}
