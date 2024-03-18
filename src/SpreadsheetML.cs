using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Address;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetML
{
    public static WorkbookReader CreateWorkbookReader(string xlsx_path) => CreateWorkbookReader(File.OpenRead(xlsx_path));

    public static WorkbookReader CreateWorkbookReader(Stream stream, bool leave_open = false) => SpreadsheetMLReader.OpenReader(stream, leave_open);

    public static Workbook Import(string xlsx_path) => Import(File.OpenRead(xlsx_path));

    public static Workbook Import(Stream stream, bool leave_open = false) => SpreadsheetMLImport.DoImport(stream, leave_open);

    public static void Export(string export_path, Workbook workbook, FormatNamespace format = FormatNamespace.Strict) => Export(export_path, workbook, SpreadsheetMLCalculation.Calculation(workbook), format);

    public static void Export(string export_path, Workbook workbook, Dictionary<string, WorksheetCalculation> calc, FormatNamespace format = FormatNamespace.Strict) => File.Create(export_path).Using(x => Export(x, workbook, calc, false, format));

    public static void Export(Stream stream, Workbook workbook, bool leave_open = false, FormatNamespace format = FormatNamespace.Strict) => Export(stream, workbook, SpreadsheetMLCalculation.Calculation(workbook), leave_open, format);

    public static void Export(Stream stream, Workbook workbook, Dictionary<string, WorksheetCalculation> calc, bool leave_open = false, FormatNamespace format = FormatNamespace.Strict) => SpreadsheetMLExport.DoExport(stream, workbook, calc, leave_open, format);


    public static CellAddress ConvertCellAddress(string cell)
    {
        var split = cell.IndexOfAny(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']);
        return new() { Row = int.Parse(cell[split..]), Column = ConvertColumnNameToIndex(cell[0..split]) };
    }

    public static CellAddressRange ConvertCellRange(string range)
    {
        var split = range.IndexOf(':');
        return new() { From = ConvertCellAddress(range[0..split]), To = ConvertCellAddress(range[(split + 1)..]) };
    }

    public static RowAddressRange ConvertRowRange(string range)
    {
        var split = range.IndexOf(':');
        return new() { From = int.Parse(range[0..split]), To = int.Parse(range[(split + 1)..]) };
    }

    public static ColumnAddressRange ConvertColumnRange(string range)
    {
        var split = range.IndexOf(':');
        return new() { From = ConvertColumnNameToIndex(range[0..split]), To = ConvertColumnNameToIndex(range[(split + 1)..]) };
    }

    public static string ConvertCellAddress(int row, int col) => $"{ConvertColumnIndexToName(col)}{row}";

    public static int ConvertColumnNameToIndex(string col) => col.Aggregate(0, (n, c) => (n * 26) + c - 'A' + 1);

    public static string ConvertColumnIndexToName(int col) => (col <= 26 ? "" : ConvertColumnIndexToName((col - 1) / 26)) + (char)('A' + ((col - 1) % 26));
}
