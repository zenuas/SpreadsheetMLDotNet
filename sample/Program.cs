﻿using Mina.Extension;
using SpreadsheetMLDotNet;
using SpreadsheetMLDotNet.Attributes;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Drawing;

var book = new Workbook();
var sheet = book.Worksheets[0];
sheet.SetCell("C2", "xC2");
sheet.SetCell("D2", "xD2");
sheet.SetCell("F2", "xF2");
sheet.SetCell("A4", "xA4");
sheet.SetCell("A6", 1_234_567);
sheet.SetCell("B6", Cell.FromFormula("(-A6+2)*(3-1)"));
sheet.SetCell("B7", Cell.FromFormula("sum(A6, B6 - 4)"));
sheet.SetCell("A7", new DateTime(2000, 1, 23, 0, 0, 0));
sheet.SetCell("A8", new DateTime(2000, 1, 23, 4, 56, 7, 89));
sheet.GetRow(3).Height = 32.1;
sheet.GetRow(3).Fill = new() { ForegroundColor = Color.Gray };
sheet.GetColumn("F").Width = 6;
sheet.GetColumn("F").BestFitColumnWidth = true;
sheet.GetColumn("F").Fill = new() { ForegroundColor = Color.LightCyan };
sheet.GetColumn("G").Width = 32.1;
sheet.GetColumn("G").Fill = new() { ForegroundColor = Color.LightCyan };
sheet.GetCell("A4").Fill = new() { ForegroundColor = IndexedColors.Indexed2.GetAttributeOrDefault<ArgbAttribute>()!.Color };
sheet.GetCell("F2").Fill = new() { ForegroundColor = IndexedColors.Indexed2.GetAttributeOrDefault<ArgbAttribute>()!.Color, PatternType = PatternTypes.DarkUp };
sheet.GetCell("C2").Border = new(Borders.Top | Borders.Bottom, BorderStyles.Thin, Color.Red);
sheet.GetCell("C2").Font = new() { Color = Color.Blue, FontSize = 22, Underline = UnderlineTypes.DoubleUnderline };
sheet.GetCell("C2").Alignment = new() { HorizontalAlignment = HorizontalAlignmentTypes.RightHorizontalAlignment };
sheet.GetCell("B6").Border = new(Borders.End, BorderStyles.Thin, null);
sheet.GetCell("A6").NumberFormat = new NumberFormatId { FormatId = NumberFormats.GeneralIntSeparate };
sheet.GetCell("A7").NumberFormat = new NumberFormatCode { FormatCode = "yyyy/mm/dd" };
sheet.GetCell("A7").Alignment = new() { ShrinkToFit = true };
sheet.GetCell("A8").NumberFormat = new NumberFormatCode { FormatCode = "yyyy/mm/dd" };
sheet.GetCell("A8").Alignment = new() { ShrinkToFit = true };
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
