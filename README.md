# SpreadsheetMLDotNet [![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

SpreadsheetMLDotNet is Office Open XML SpreadsheetML library  

## Description

SpreadsheetMLDotNet is a poor sample program.  

SpreadsheetMLDotNetは貧弱なサンプルプログラムである。  
簡易の読み込みと書き込みをサポートする。  

## Usage

WorkbookReaderはセルの値のみ取得する。
式の評価やスタイル、フォーマットは取得しない。

```cs
using var workbook = SpreadsheetML.CreateWorkbookReader("Test.xlsx");
foreach (var name in workbook.WorkSheetNames)
{
    Console.WriteLine(name);
    using var worksheet = workbook.OpenWorksheet(name);
    foreach (var (cell, value) in worksheet)
    {
        Console.WriteLine($"{cell} = {value}");
    }
}
```

Workbookは一部の編集とエクスポートを行う。

```cs
var book = new Workbook();
var sheet = book.Worksheets[0];
sheet.SetCell("A1", "cell text");
SpreadsheetML.Export("Test.xlsx", book);
```
