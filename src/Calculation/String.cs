namespace SpreadsheetMLDotNet.Calculation;

public readonly struct String : IFormula
{
    public required string Value { get; init; }
}
