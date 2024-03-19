using SpreadsheetMLDotNet;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Drawing;

var book = new Workbook();
var sheet = book.Worksheets[0];
for (var col = 1; col < 10; col++)
{
    sheet.SetCell(1, col, $"Field{SpreadsheetML.ConvertColumnIndexToName(col)}");
    for (var row = 2; row < 102; row++)
    {
        if (col <= 2)
        {
            sheet.SetCell(row, col, (row - 1) * col);
        }
        else
        {
            var addr = SpreadsheetML.ConvertCellAddress(row, col);
            sheet.SetCell(row, col, $"x{addr}");
        }
    }
}
sheet.SetCell("A102", Cell.FromFormula("SUM(A2:A101)"));
sheet.GetCell("A102").Fill = new() { ForegroundColor = Color.Yellow };
SpreadsheetML.Export("New.Strict.xlsx", book, FormatNamespace.Strict);
SpreadsheetML.Export("New.Transitional.xlsx", book, FormatNamespace.Transitional);

using var workbook = SpreadsheetML.CreateWorkbookReader("New.Transitional.xlsx");
foreach (var name in workbook.WorkSheetNames)
{
    Console.WriteLine(name);
    using var worksheet = workbook.OpenWorksheet(name);
    foreach (var (cell, value) in worksheet)
    {
        Console.WriteLine($"{cell} = {value}");
    }
}

var workbook2 = SpreadsheetML.Import("New.Strict.xlsx");
SpreadsheetML.Export("New.Strict2.xlsx", workbook2, FormatNamespace.Strict);
