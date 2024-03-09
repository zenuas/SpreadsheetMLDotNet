using System;

namespace SpreadsheetMLDotNet.Calculation;

public class Date : IFormula
{
    public required DateTime Value { get; init; }
}
