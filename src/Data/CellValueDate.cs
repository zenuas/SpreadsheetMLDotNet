using System;

namespace SpreadsheetMLDotNet.Data;

public class CellValueDate : ICellValue
{
    public required DateTime Value { get; set; }
}
