using Mina.Extension;
using SpreadsheetMLDotNet.Calculation;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLCalculation
{
    public static Dictionary<string, WorksheetCalculation> Calculation(Workbook workbook)
    {
        var calc = workbook.Worksheets.ToDictionary(x => x.Name, x => new WorksheetCalculation { Worksheet = x });
        foreach (var (y, x, cell, worksheet) in SpreadsheetMLExport.EnumerableCells(workbook).Where(x => x.Cell.Value is CellValueFormula))
        {
            CalculationCell(calc, worksheet, cell.Value.Cast<CellValueFormula>(), SpreadsheetML.ConvertCellAddress(y, x));
        }
        return calc;
    }

    public static ICellValue CalculationCell(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, string addr)
    {
        var cell = current_sheet.GetCellOrDefault(addr);
        return cell is null ? CellValueNull.Instance
            : cell.Value is CellValueFormula formula ? CalculationCell(calc, current_sheet, formula, addr)
            : cell.Value;
    }

    public static ICellValue CalculationCell(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, CellValueFormula formula, string addr)
    {
        var calcwork = calc[current_sheet.Name].Calculation;
        if (!calcwork.TryGetValue(addr, out var value))
        {
            calcwork.Add(addr, null);
            return calcwork[addr] = Evaluate(calc, current_sheet, Parse(formula.Value));
        }
        else
        {
            return value is { } ? value : throw new("circular reference");
        }
    }

    public static ICellValue Evaluate(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, IFormula formula)
    {
        switch (formula)
        {
            case Token token:
                return CalculationCell(calc, current_sheet, token.Value);

            case Unary unary:
                if (unary.Operator.In("+", "()")) return Evaluate(calc, current_sheet, unary.Value);
                if (unary.Operator == "-")
                {
                    var value = Evaluate(calc, current_sheet, unary.Value);
                    if (value is CellValueDouble d) return new CellValueDouble { Value = -d.Value };
                }
                return CellValueError.VALUE;

            case Expression expr:
                var left = Evaluate(calc, current_sheet, expr.Left);
                var right = Evaluate(calc, current_sheet, expr.Right);
                switch (left)
                {
                    case CellValueString str:
                        return EvaluateString(str, expr.Operator, right);

                    case CellValueDouble num:
                        return EvaluateDouble(num, expr.Operator, right);

                    case CellValueDate date:
                        return EvaluateDate(date, expr.Operator, right);
                }
                break;

            case Number number:
                return new CellValueDouble { Value = number.Value };

            case Calculation.String str:
                return new CellValueString { Value = str.Value };

            case Date date:
                return new CellValueDate { Value = date.Value };

            case FunctionCall call:
                return EvaluateFunction(calc, current_sheet, call);

            case ICellValue cell:
                return cell;
        }
        return CellValueNull.Instance;
    }

    public static ICellValue EvaluateDouble(CellValueDouble left, string op, ICellValue right)
    {
        var rx = EvaluateToDouble(right);
        return rx is null ? CellValueError.VALUE
            : op switch
            {
                "+" => new CellValueDouble { Value = left.Value + rx.Value },
                "-" => new CellValueDouble { Value = left.Value - rx.Value },
                "*" => new CellValueDouble { Value = left.Value * rx.Value },
                "/" => rx.Value == 0 ? CellValueError.DIV_0 : new CellValueDouble { Value = left.Value / rx.Value },
                "=" => new CellValueBoolean { Value = left.Value.Equals(rx.Value) },
                _ => CellValueError.VALUE,
            };
    }

    public static ICellValue EvaluateString(CellValueString left, string op, ICellValue right)
    {
        var rx = EvaluateToString(right);
        return op switch
        {
            "+" => new CellValueString { Value = left.Value + rx },
            "&" => new CellValueString { Value = left.Value + rx },
            "=" => new CellValueBoolean { Value = left.Value == rx },
            _ => CellValueError.VALUE,
        };
    }

    public static ICellValue EvaluateDate(CellValueDate left, string op, ICellValue right)
    {
        if (op == "=") return new CellValueBoolean { Value = left.Value.Equals(EvaluateToDate(right)) };
        var rx = EvaluateToDouble(right);
        return rx is null ? CellValueError.VALUE
            : op switch
            {
                "+" => new CellValueDate { Value = left.Value.AddDays(rx.Value) },
                _ => CellValueError.VALUE,
            };
    }

    public static double? EvaluateToDouble(ICellValue value) => value switch
    {
        CellValueString x => double.TryParse(x.Value, out var d) ? d : null,
        CellValueDouble x => x.Value,
        CellValueBoolean x => x.Value ? 1 : 0,
        _ => null
    };

    public static string EvaluateToString(ICellValue value) => value switch
    {
        CellValueString x => x.Value,
        CellValueDouble x => x.Value.ToString(),
        CellValueDate x => x.Value.ToString(),
        CellValueBoolean x => x.Value ? "TRUE" : "FALSE",
        _ => ""
    };

    public static DateTime? EvaluateToDate(ICellValue value) => value switch
    {
        CellValueDate x => x.Value,
        _ => null
    };
}
