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
        var cell = current_sheet.TryGetCell(addr);
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

    public static IFormula Parse(string formula) => Parse(ParseTokens(formula), 0).Formula;

    public static (IFormula Formula, int Length) Parse(Span<(TokenTypes Type, string Value)> values, int parenthesis_level)
    {
        var (left, next) = ParseValue(values, parenthesis_level);
        if (next >= values.Length) return (left, next);

        switch (values[next].Type)
        {
            case TokenTypes.Operator:
                {
                    var (right, length) = Parse(values[(next + 1)..], parenthesis_level);
                    return right is Expression rx && rx.Operator.In("+", "-") && values[next].Value.In("*", "/")
                        ? (new Expression { Operator = rx.Operator, Left = new Expression { Operator = values[next].Value, Left = left, Right = rx.Left }, Right = rx.Right }, next + length + 1)
                        : (new Expression { Operator = values[next].Value, Left = left, Right = right }, next + length + 1);
                }

            case TokenTypes.LeftParenthesis:
                {
                    var fname =
                        left is Token token ? token.Value
                        : left is Unary lu && lu.Operator != "()" && lu.Value is Token unary_token ? unary_token.Value
                        : "";
                    if (fname == "") break;
                    var args = new List<IFormula>();
                    var total_length = next + 1;
                    while (true)
                    {
                        var (arg, length) = Parse(values[total_length..], parenthesis_level + 1);
                        if (arg is Error || arg is Null) break;
                        args.Add(arg);
                        total_length += length;
                    }
                    var fcall = new FunctionCall { Name = fname.ToUpperInvariant(), Arguments = [.. args] };
                    return (left is Unary unary ? new Unary { Operator = unary.Operator, Value = fcall } : fcall, total_length);
                }

            case TokenTypes.Comma:
                return (left, next + 1);

            case TokenTypes.RightParenthesis:
                return (left, next);
        }
        return (new Error(), 0);
    }

    public static (IFormula Formula, int Length) ParseValue(Span<(TokenTypes Type, string Value)> values, int parenthesis_level)
    {
        if (values.Length == 0) return (new Null(), 0);

        if (values[0].Type == TokenTypes.LeftParenthesis)
        {
            var (value, next) = Parse(values[1..], parenthesis_level + 1);
            return (new Unary { Operator = "()", Value = value }, next + 2);
        }
        else if (values[0].Type == TokenTypes.RightParenthesis)
        {
            return parenthesis_level <= 0 ? (new Error(), 0) : (new Null(), 1);
        }
        else if (values[0].Type == TokenTypes.Operator)
        {
            var (value, next) = ParseValue(values[1..], parenthesis_level);
            return (new Unary { Operator = values[0].Value, Value = value }, next + 1);
        }
        else
        {
            return (ParsePrimitive(values[0].Type, values[0].Value), 1);
        }
    }

    public static IFormula ParsePrimitive(TokenTypes type, string value) =>
        type == TokenTypes.Token ? new Token { Value = value } :
        type == TokenTypes.String ? new Token { Value = value } :
        type == TokenTypes.Number ? new Number { Value = double.Parse(value) } :
        new Error();

    public static (TokenTypes Type, string Value)[] ParseTokens(string formula)
    {
        var tokens = new List<(TokenTypes Type, string Value)>();
        for (var i = 0; i < formula.Length; i++)
        {
            var c = formula[i];
            if (char.IsWhiteSpace(formula[i])) continue;
            switch (c)
            {
                case '(':
                    tokens.Add((TokenTypes.LeftParenthesis, "("));
                    break;

                case ')':
                    tokens.Add((TokenTypes.RightParenthesis, ")"));
                    break;

                case ',':
                    tokens.Add((TokenTypes.Comma, ","));
                    break;

                case ':':
                    tokens.Add((TokenTypes.Range, ":"));
                    break;

                case '"':
                    {
                        var (value, length) = ParseString(formula[i..]);
                        tokens.Add((TokenTypes.String, value));
                        i += length - 1;
                    }
                    break;

                default:
                    if (char.IsAsciiDigit(c))
                    {
                        var (value, length) = ParseNumber(formula[i..]);
                        tokens.Add((TokenTypes.Number, value));
                        i += length - 1;
                    }
                    else if (char.IsAsciiLetter(c))
                    {
                        var (value, length) = ParseToken(formula[i..]);
                        tokens.Add((TokenTypes.Token, value));
                        i += length - 1;
                    }
                    else if (IsOperator(c))
                    {
                        var (value, length) = ParseOperator(formula[i..]);
                        tokens.Add((TokenTypes.Operator, value));
                        i += length - 1;
                    }
                    break;
            }
        }
        return [.. tokens];
    }

    public static (string Value, int Length) ParseString(string s)
    {
        if (s[0] != '"') return ("", 0);
        int i = 1;
        for (; i < s.Length; i++)
        {
            if (s[i] == '"')
            {
                if (i + 1 >= s.Length || s[i + 1] != '"') break;
                i++;
            }
        }
        return (s[1..i].Replace("\"\"", "\""), i + 1);
    }

    public static (string Value, int Length) ParseNumber(string s)
    {
        var i = 0;
        for (; i < s.Length; i++)
        {
            if (!char.IsAsciiDigit(s[i])) break;
        }
        return (s[0..i], i);
    }

    public static (string Value, int Length) ParseToken(string s)
    {
        var i = 0;
        for (; i < s.Length; i++)
        {
            if (!IsWord(s[i])) break;
        }
        return (s[0..i], i);
    }

    public static (string Value, int Length) ParseOperator(string s)
    {
        var c = s[0];
        return (c == '<' || c == '>') && 1 < s.Length && s[1] == '='
            ? (s[0..2], 2)
            : (s[0..1], 1);
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

    public static bool IsWord(char c) => c == '_' || char.IsAsciiLetterOrDigit(c);

    public static bool IsOperator(char c) =>
        c == '-' ||
        c == '+' ||
        c == '*' ||
        c == '/' ||
        c == '&' ||
        c == '=' ||
        c == '<' ||
        c == '>';
}
