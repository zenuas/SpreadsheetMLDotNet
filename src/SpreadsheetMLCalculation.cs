﻿using Mina.Extension;
using SpreadsheetMLDotNet.Calculation;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLCalculation
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
            return calcwork[addr] = Evaluate(Parse(formula.Value), calc, current_sheet);
        }
        else
        {
            return value is { } ? value : throw new("circular reference");
        }
    }

    public static IFormula Parse(string formula) => Parse(ParseTokens(formula));

    public static IFormula Parse(Span<(TokenTypes Type, string Value)> values)
    {
        if (values.Length == 0) return new Null();
        if (values.Length == 1) return ParseValue(values[0].Type, values[0].Value);

        if (values[0].Type == TokenTypes.LeftParenthesis)
        {

        }
        else
        {
            switch (values[1].Type)
            {
                case TokenTypes.Operator:
                    var left = ParseValue(values[0].Type, values[0].Value);
                    var right = Parse(values[2..]);
                    if (right is Expression rx && rx.Operator.In("+", "-") && values[1].Value.In("*", "/"))
                    {
                        return new Expression() { Operator = rx.Operator, Left = new Expression() { Operator = values[1].Value, Left = left, Right = rx.Left }, Right = rx.Right };
                    }
                    return new Expression() { Operator = values[1].Value, Left = left, Right = right };

                case TokenTypes.LeftParenthesis:
                    break;
            }
        }
        return new Error();
    }

    public static IFormula ParseValue(TokenTypes type, string value) =>
        type == TokenTypes.Token ? new Token() { Value = value } :
        type == TokenTypes.String ? new Token() { Value = value } :
        type == TokenTypes.Number ? new Number() { Value = double.Parse(value) } :
        throw new();

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
                        var j = i + 1;
                        for (; j < formula.Length; j++)
                        {
                            if (formula[j] == '"')
                            {
                                if (j + 1 >= formula.Length || formula[j + 1] != '"') break;
                                j++;
                            }
                        }
                        tokens.Add((TokenTypes.String, formula[(i + 1)..j].Replace("\"\"", "\"")));
                        i = j;
                    }
                    break;

                default:
                    if (char.IsAsciiDigit(c))
                    {
                        var j = i + 1;
                        for (; j < formula.Length; j++)
                        {
                            if (!char.IsAsciiDigit(formula[j])) break;
                        }
                        tokens.Add((TokenTypes.Number, formula[i..j]));
                        i = j - 1;
                    }
                    else if (char.IsAsciiLetter(c))
                    {
                        var j = i + 1;
                        for (; j < formula.Length; j++)
                        {
                            if (!IsWord(formula[j])) break;
                        }
                        tokens.Add((TokenTypes.Token, formula[i..j]));
                        i = j - 1;
                    }
                    else if (IsOperator(c))
                    {
                        if ((c == '<' || c == '>') && i + 1 < formula.Length && formula[i + 1] == '=')
                        {
                            tokens.Add((TokenTypes.Operator, formula[i..(i + 2)]));
                            i++;
                        }
                        else
                        {
                            tokens.Add((TokenTypes.Operator, c.ToString()));
                        }
                    }
                    break;
            }
        }
        return [.. tokens];
    }

    public static ICellValue Evaluate(IFormula formula, Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet)
    {
        switch (formula)
        {
            case Unary unary:
                break;

            case Expression expr:
                break;
        }
        return CellValueNull.Instance;
    }

    public static bool IsWord(char c) => c == '_' || char.IsAsciiLetterOrDigit(c);

    public static bool IsOperator(char c) =>
        c == '-' ||
        c == '+' ||
        c == '*' ||
        c == '/' ||
        c == '<' ||
        c == '>';
}