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
SpreadsheetML.Export("New.xlsx", book);
