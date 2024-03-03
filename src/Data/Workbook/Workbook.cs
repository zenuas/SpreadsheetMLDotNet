using SpreadsheetMLDotNet.Data.Worksheets;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.Workbook;

public class Workbook : IRelationshipable
{
    public List<Worksheet> Worksheets { get; init; } = [new() { Name = "Sheet1" }];
}
