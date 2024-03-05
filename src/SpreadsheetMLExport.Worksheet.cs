﻿using Mina.Attributes;
using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
{
    public static void WriteWorksheet(Stream stream, Worksheet worksheet, FormatNamespace format, List<CellStyle> cellstyles)
    {
        stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""{FormatNamespaces.SpreadsheetMLMains[(int)format]}"">
  <sheetData>
");
        foreach (var (y, row) in EnumerableRows(worksheet))
        {
            var row_attr = new Dictionary<string, string>();
            if (TryAddAttribute(row_attr, "ht", row.Height)) row_attr["customHeight"] = "1";
            if (row.Cells.Count == 0 && row_attr.Count == 0) continue;

            row_attr["r"] = y.ToString();

            stream.Write($"    <row {AttributesToString(row_attr)}>\r\n");
            foreach (var (x, cell) in EnumerableCells(row))
            {
                var cell_attr = new Dictionary<string, string>();
                if (TryAddStyleIndex(cell, cellstyles, out var styleindex)) cell_attr["s"] = styleindex.ToString();
                if (cell.Value is CellValueNull && cell_attr.Count == 0) continue;

                var (cell_type, escaped_value) = GetCellValueFormat(cell.Value);
                cell_attr["r"] = SpreadsheetML.ConvertCellAddress(y, x);
                cell_attr["t"] = cell_type.GetAttributeOrDefault<AliasAttribute>()!.Name;

                stream.Write(
$@"      <c {AttributesToString(cell_attr)}>
        <v>{escaped_value}</v>
      </c>
");
            }
            stream.Write($"    </row>\r\n");
        }
        stream.Write(
$@"  </sheetData>
</worksheet>
");
    }
    public static IEnumerable<(int Index, Row Row)> EnumerableRows(Worksheet worksheet)
    {
        for (var y = 0; y < worksheet.Rows.Count; y++)
        {
            yield return (y + worksheet.StartRowIndex, worksheet.Rows[y]);
        }
    }

    public static IEnumerable<(int Index, Cell Cell)> EnumerableCells(Row row)
    {
        for (var x = 0; x < row.Cells.Count; x++)
        {
            yield return (x + row.StartCellIndex, row.Cells[x]);
        }
    }

    public static IEnumerable<(int Row, int Column, Cell Cell)> EnumerableCells(Workbook workbook)
    {
        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var (y, row) in EnumerableRows(worksheet))
            {
                foreach (var (x, cell) in EnumerableCells(row))
                {
                    yield return (y, x, cell);
                }
            }
        }
    }

    public static (CellTypes CellType, string EscapedValue) GetCellValueFormat(ICellValue value) => value switch
    {
        CellValueBoolean x => (CellTypes.Boolean, x.Value.ToString()),
        CellValueDate x => (CellTypes.Date, x.Value.ToString()),
        CellValueError x => (CellTypes.Error, x.Value.GetAttributeOrDefault<AliasAttribute>()!.Name),
        CellValueDouble x => (CellTypes.Number, x.Value.ToString()),
        CellValueString x => (CellTypes.String, SecurityElement.Escape(x.Value)),
        CellValueNull => (CellTypes.String, ""),
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };
}