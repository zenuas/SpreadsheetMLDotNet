using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Calculation;

public class FunctionCall : IFormula
{
    public required string Name { get; init; }
    public List<IFormula> Arguments { get; init; } = [];
}
