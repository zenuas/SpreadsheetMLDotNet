using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data;

public class Workbook : IRelationshipable
{
    public List<Worksheet> Worksheets { get; init; } = [new() { Name = "Sheet1" }];
}
