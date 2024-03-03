using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Workbook;
using System.IO;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetML
{
    public static WorkbookReader CreateWorkbookReader(string xlsx_path) => CreateWorkbookReader(File.OpenRead(xlsx_path));

    public static WorkbookReader CreateWorkbookReader(Stream stream, bool leave_open = false) => SpreadsheetMLReader.OpenReader(stream, leave_open);

    public static void Export(string export_path, Workbook workbook, FormatNamespace format = FormatNamespace.Strict) => File.Create(export_path).Using(x => Export(x, workbook, false, format));

    public static void Export(Stream stream, Workbook workbook, bool leave_open = false, FormatNamespace format = FormatNamespace.Strict) => SpreadsheetMLExport.DoExport(stream, workbook, leave_open, format);


    public static (int Row, int Column) ConvertCellAddress(string cell)
    {
        var split = cell.IndexOfAny(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']);
        return (int.Parse(cell[split..]), ConvertColumnNameToIndex(cell[0..split]));
    }

    public static string ConvertCellAddress(int row, int col) => $"{ConvertColumnIndexToName(col)}{row}";

    public static int ConvertColumnNameToIndex(string col) => col.Aggregate(0, (n, c) => (n * 26) + c - 'A' + 1);

    public static string ConvertColumnIndexToName(int col) => (col <= 26 ? "" : ConvertColumnIndexToName((col - 1) / 26)) + (char)('A' + ((col - 1) % 26));
}
