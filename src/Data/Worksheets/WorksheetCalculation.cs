using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class WorksheetCalculation
{
    public required Worksheet Worksheet { get; init; }
    public Dictionary<string, ICellValue?> Calculation { get; init; } = [];
}
