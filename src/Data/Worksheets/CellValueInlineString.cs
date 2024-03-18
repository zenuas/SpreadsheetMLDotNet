using SpreadsheetMLDotNet.Data.Styles;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class CellValueInlineString : ICellValue
{
    public List<RichText> Values { get; init; } = [];
}
