using Mina.Extension;
using SpreadsheetMLDotNet;
using SpreadsheetMLDotNet.Attributes;
using SpreadsheetMLDotNet.Data;
using System;

var book = new Workbook();
var sheet = book.Worksheets[0];
sheet.SetCell("C2", "xC2");
sheet.SetCell("D2", "xD2");
sheet.SetCell("F2", "xF2");
sheet.SetCell("A4", "xA4");
sheet.GetRow(3).Height = 32.1;
sheet.GetCell("A4").Fill = new() { ForegroundColor = IndexedColors.Indexed2.GetAttributeOrDefault<ArgbAttribute>()!.Color };
sheet.GetCell("F2").Fill = new() { ForegroundColor = IndexedColors.Indexed2.GetAttributeOrDefault<ArgbAttribute>()!.Color, PatternType = PatternTypes.DarkUp };
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
