using Mina.Attributes;
using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
{
    public static void WriteWorksheet(Stream stream, Worksheet worksheet, FormatNamespace format, List<CellStyle> cellstyles, Dictionary<string, WorksheetCalculation> calc)
    {
        stream.WriteLine($"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{FormatNamespaces.SpreadsheetMLMains[(int)format]}">
""");
        if (worksheet.Columns.Count > 0)
        {
            stream.WriteLine("  <cols>");
            foreach (var (x, col) in EnumerableColumns(worksheet))
            {
                var col_attr = new Dictionary<string, string>();
                TryAddAttribute(col_attr, "hidden", col.Hidden);
                TryAddAttribute(col_attr, "outlineLevel", col.OutlineLevel);
                TryAddAttribute(col_attr, "collapsed", col.Collapsed);
                TryAddAttribute(col_attr, "phonetic", col.ShowPhonetic);
                if (TryAddAttribute(col_attr, "width", col.Width)) col_attr["customWidth"] = "1";
                if (TryAddAttribute(col_attr, "bestFit", col.BestFitColumnWidth)) col_attr["customWidth"] = "1";
                if (TryAddStyleIndex(col, cellstyles, out var col_styleindex)) col_attr["style"] = col_styleindex.ToString();
                if (col_attr.Count == 0) continue;

                col_attr["min"] = col_attr["max"] = x.ToString();
                stream.WriteLine($"    <col {AttributesToString(col_attr)}/>");
            }
            stream.WriteLine("  </cols>");
        }

        stream.WriteLine($"  <sheetData{(worksheet.Rows.Count == 0 ? "/" : "")}>");
        foreach (var (y, row) in EnumerableRows(worksheet))
        {
            var row_attr = new Dictionary<string, string>();
            TryAddAttribute(row_attr, "hidden", row.Hidden);
            TryAddAttribute(row_attr, "outlineLevel", row.OutlineLevel);
            TryAddAttribute(row_attr, "collapsed", row.Collapsed);
            TryAddAttribute(row_attr, "thickTop", row.ThickTop);
            TryAddAttribute(row_attr, "thickBot", row.ThickBottom);
            TryAddAttribute(row_attr, "ph", row.ShowPhonetic);
            if (TryAddAttribute(row_attr, "ht", row.Height)) row_attr["customHeight"] = "1";
            if (TryAddStyleIndex(row, cellstyles, out var row_styleindex)) { row_attr["s"] = row_styleindex.ToString(); row_attr["customFormat"] = "1"; }
            if (row.Cells.Count == 0 && row_attr.Count == 0) continue;

            row_attr["r"] = y.ToString();

            stream.WriteLine($"    <row {AttributesToString(row_attr)}{(row.Cells.Count == 0 ? "/" : "")}>");
            foreach (var (x, cell) in EnumerableCells(row))
            {
                var cell_attr = new Dictionary<string, string>();
                if (TryAddStyleIndex(cell, cellstyles, out var cell_styleindex)) cell_attr["s"] = cell_styleindex.ToString();
                if (cell.Value is CellValueNull && cell_attr.Count == 0) continue;

                var addr = cell_attr["r"] = SpreadsheetML.ConvertCellAddress(y, x);
                var (cell_type, escaped_value) = GetCellValueFormat(cell.Value is CellValueFormula && calc.TryGetValue(worksheet.Name, out var xc) && xc.Calculation.TryGetValue(addr, out var xv) ? xv! : cell.Value);
                cell_attr["t"] = cell_type.GetAttributeOrDefault<AliasAttribute>()!.Name;

                stream.WriteLine($"      <c {AttributesToString(cell_attr)}{(escaped_value == "" && cell.Value is not CellValueFormula && (cell_type != CellTypes.InlineString || cell.Value.Cast<CellValueInlineString>().Values.Count == 0) ? "/" : "")}>");
                if (cell_type == CellTypes.InlineString && cell.Value is CellValueInlineString instr && instr.Values.Count > 0)
                {
                    stream.WriteLine("        <is>");
                    foreach (var rt in instr.Values)
                    {
                        stream.WriteLine("          <r>");
                        if (ExistsFontSetting(rt))
                        {
                            stream.WriteLine("            <rPr>");
                            WriteFontAttribute(stream, 14, rt);
                            stream.WriteLine("            </rPr>");
                        }
                        stream.WriteLine($"            <t>{SecurityElement.Escape(rt.Text)}</t>");
                        stream.WriteLine("          </r>");
                    }
                    stream.WriteLine("        </is>");
                    stream.WriteLine("      </c>");
                }
                else if (escaped_value != "" || cell.Value is CellValueFormula)
                {
                    if (cell.Value is CellValueFormula formula) stream.WriteLine($"        <f>{formula.Value}</f>");
                    if (escaped_value != "") stream.WriteLine($"        <v>{escaped_value}</v>");
                    stream.WriteLine("      </c>");
                }
            }
            if (row.Cells.Count > 0) stream.WriteLine("    </row>");
        }
        if (worksheet.Rows.Count > 0) stream.WriteLine("  </sheetData>");
        if (worksheet.AutoFilter is { } af)
        {
            stream.WriteLine($"""  <autoFilter ref="{af.Reference.ToAddressName()}"/>""");
        }
        if (worksheet.Merges.Count > 0)
        {
            stream.WriteLine($"""  <mergeCells count="{worksheet.Merges.Count}">""");
            foreach (var merge in worksheet.Merges)
            {
                stream.WriteLine($"""    <mergeCell ref="{merge.ToAddressName()}"/>""");
            }
            stream.WriteLine("  </mergeCells>");
        }
        stream.WriteLine("</worksheet>");
    }

    public static IEnumerable<(int Index, Row Row)> EnumerableRows(Worksheet worksheet) => worksheet.Rows.GetEnumerableWithIndex();

    public static IEnumerable<(int Index, Cell Cell)> EnumerableCells(Row row) => row.Cells.GetEnumerableWithIndex();

    public static IEnumerable<(int Row, int Column, Cell Cell, Worksheet Worksheet)> EnumerableCells(Workbook workbook)
    {
        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var (y, x, cell) in EnumerableCells(worksheet))
            {
                yield return (y, x, cell, worksheet);
            }
        }
    }

    public static IEnumerable<(int Row, int Column, Cell Cell)> EnumerableCells(Worksheet worksheet)
    {
        foreach (var (y, row) in EnumerableRows(worksheet))
        {
            foreach (var (x, cell) in EnumerableCells(row))
            {
                yield return (y, x, cell);
            }
        }
    }

    public static IEnumerable<(int Index, Column Column)> EnumerableColumns(Worksheet worksheet) => worksheet.Columns.GetEnumerableWithIndex();

    public static (CellTypes CellType, string EscapedValue) GetCellValueFormat(ICellValue value) => value switch
    {
        CellValueBoolean x => (CellTypes.Boolean, x.Value.ToString()),
        CellValueDate x => (CellTypes.Date, x.Value.ToString(x.Value.Ticks % TimeSpan.TicksPerDay == 0 ? "yyyy-MM-dd" : "o")), // ISO 8601 style, if date is short style in Excel
        CellValueError x => (CellTypes.Error, x.Value.GetAttributeOrDefault<AliasAttribute>()!.Name),
        CellValueDouble x => (CellTypes.Number, x.Value.ToString()),
        CellValueString x => (CellTypes.String, SecurityElement.Escape(x.Value)),
        CellValueNull => (CellTypes.String, ""),
        CellValueFormula => (CellTypes.String, ""),
        CellValueInlineString => (CellTypes.InlineString, ""),
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };
}
