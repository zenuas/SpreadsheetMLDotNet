using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Calculation;

public readonly struct FunctionCall() : IFormula
{
    public required string Name { get; init; }
    public List<IFormula> Arguments { get; init; } = [];
}
