using System;

namespace SpreadsheetMLDotNet.Calculation;

public readonly struct Date : IFormula
{
    public required DateTime Value { get; init; }
}
