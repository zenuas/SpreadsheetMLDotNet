using System;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class CellValueDate : ICellValue
{
    public required DateTime Value { get; set; }
}
