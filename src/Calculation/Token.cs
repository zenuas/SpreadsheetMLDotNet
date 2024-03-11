namespace SpreadsheetMLDotNet.Calculation;

public readonly struct Token : IFormula
{
    public required string Value { get; init; }
}
