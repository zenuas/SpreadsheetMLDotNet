namespace SpreadsheetMLDotNet.Calculation;

public readonly struct FunctionCall() : IFormula
{
    public required string Name { get; init; }
    public IFormula[] Arguments { get; init; } = [];
}
