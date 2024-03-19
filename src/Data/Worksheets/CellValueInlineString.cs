using SpreadsheetMLDotNet.Data.SharedStringTable;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class CellValueInlineString : ICellValue
{
    public RunProperties Values { get; init; } = [];
}
