using Mina.Extension;
using SpreadsheetMLDotNet.Calculation;
using SpreadsheetMLDotNet.Data.Address;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLCalculation
{
    public static ICellValue EvaluateFunction(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, FunctionCall call) => call.Name switch
    {
        "SUM" => Sum(calc, current_sheet, call.Arguments),
        _ => CellValueError.NAME,
    };

    public static ICellValue Sum(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, Span<IFormula> values)
    {
        if (values.Length == 0) return new CellValueDouble { Value = 0 };

        CellValueDouble r;
        if (values[0] is Calculation.Range range)
        {
            r = new CellValueDouble { Value = 0 };
            foreach (var e in EnumerableCells(GetWorksheet(calc, current_sheet, range), range.From, range.To))
            {
                var x = EvaluateDouble(r, "+", e.Cell);
                if (x is not CellValueDouble r2) return CellValueError.VALUE;
                r = r2;
            }
        }
        else
        {
            var v = Evaluate(calc, current_sheet, values[0]);
            if (v is not CellValueDouble r2) return CellValueError.VALUE;
            r = r2;
        }
        return values.Length == 1 ? r : EvaluateDouble(r, "+", Sum(calc, current_sheet, values[1..]));

    }

    public static Worksheet GetWorksheet(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, Calculation.Range range) => range.SheetName == "" ? current_sheet : calc[range.SheetName].Worksheet;

    public static IEnumerable<(int Row, int Column, ICellValue Cell)> EnumerableCells(Worksheet sheet, IAddress from, IAddress to)
    {
        if (from is RowAddress r)
        {
            var start = r.Row;
            var end = to.Cast<RowAddress>().Row;
            return EnumerableRowsToCell(sheet, Math.Min(start, end), Math.Max(start, end));
        }
        if (from is ColumnAddress c)
        {
            var start = c.Column;
            var end = to.Cast<ColumnAddress>().Column;
            return EnumerableColumnsToCell(sheet, Math.Min(start, end), Math.Max(start, end));
        }
        else
        {
            var start = from.Cast<CellAddress>();
            var end = to.Cast<CellAddress>();
            return EnumerableRangeToCell(sheet, Math.Min(start.Row, end.Row), Math.Max(start.Row, end.Row), Math.Min(start.Column, end.Column), Math.Max(start.Column, end.Column));
        }
    }

    public static IEnumerable<(int Row, int Column, ICellValue Cell)> EnumerableRowsToCell(Worksheet sheet, int from, int to)
    {
        foreach (var (row, row_value) in sheet.GetRowsExists(from, to))
        {
            for (var column = 0; column < row_value.Cells.Count; column++)
            {
                yield return (row, column + row_value.Cells.StartIndex, row_value.Cells[column].Value);
            }
        }
    }

    public static IEnumerable<(int Row, int Column, ICellValue Cell)> EnumerableColumnsToCell(Worksheet sheet, int from, int to)
    {
        for (var row = 0; row < sheet.Rows.Count; row++)
        {
            var row_value = sheet.Rows[row];
            for (var column = Math.Max(0, from - row_value.Cells.StartIndex); column < row_value.Cells.Count; column++)
            {
                if (column + row_value.Cells.StartIndex > to) break;
                yield return (row + sheet.Rows.StartIndex, column + row_value.Cells.StartIndex, row_value.Cells[column].Value);
            }
        }
    }

    public static IEnumerable<(int Row, int Column, ICellValue Cell)> EnumerableRangeToCell(Worksheet sheet, int row_from, int row_to, int col_from, int col_to)
    {
        foreach (var (row, row_value) in sheet.GetRowsExists(row_from, row_to))
        {
            for (var column = Math.Max(0, col_from - row_value.Cells.StartIndex); column < row_value.Cells.Count; column++)
            {
                if (column + row_value.Cells.StartIndex > col_to) break;
                yield return (row, column + row_value.Cells.StartIndex, row_value.Cells[column].Value);
            }
        }
    }
}
