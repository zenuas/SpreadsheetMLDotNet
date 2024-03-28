using SpreadsheetMLDotNet.Data.Address;

namespace SpreadsheetMLDotNet.Calculation;

public readonly struct Range() : IFormula
{
    public string SheetName { get; init; } = "";
    public required IAddress From { get; init; }
    public required IAddress To { get; init; }
}
