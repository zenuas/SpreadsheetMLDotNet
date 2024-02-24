using SpreadsheetMLDotNet;
using SpreadsheetMLDotNet.Data;
using System;

using var workbook = SpreadsheetML.CreateWorkbookReader("Calc.xlsx");
foreach (var name in workbook.WorkSheetNames)
{
    Console.WriteLine(name);
    using var worksheet = workbook.OpenWorksheet(name);
    foreach (var (cell, value) in worksheet)
    {
        Console.WriteLine($"{cell} = {value}");
    }
}

var book = new Workbook();
var sheet = book.Worksheets[0];
sheet.SetRow(2, new Row());
var row2 = sheet.GetRow(2)!;
row2.SetCell(3, new Cell { Value = new CellValueString { Value = "xC2" } });
row2.SetCell(4, new Cell { Value = new CellValueString { Value = "xD2" } });
row2.SetCell(6, new Cell { Value = new CellValueString { Value = "xF2" } });
sheet.SetRow(4, new Row());
var row4 = sheet.GetRow(4)!;
row4.SetCell(1, new Cell { Value = new CellValueString { Value = "xA4" } });
SpreadsheetML.Export("New.Strict.xlsx", book, FormatNamespace.Strict);
SpreadsheetML.Export("New.Transitional.xlsx", book, FormatNamespace.Transitional);
