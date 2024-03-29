﻿using Mina.Extension;
using SpreadsheetMLDotNet.Calculation;
using System;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLParser
{
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
            return values[0].Type == TokenTypes.String
                ? (ParsePrimitive(values[0].Type, values[0].Value), 1)
                : values.Length >= 3 && values[1].Type == TokenTypes.Range && values[0].Type == values[2].Type
                ? (new Calculation.Range { From = SpreadsheetML.ConvertAnyAddress(values[0].Value), To = SpreadsheetML.ConvertAnyAddress(values[2].Value) }, 3)
                : (ParsePrimitive(values[0].Type, values[0].Value), 1);
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
        return !IsOperator(c) ? ("", 0) :
            ((c == '<' || c == '>') && 1 < s.Length && s[1] == '='
            ? (s[0..2], 2)
            : (s[0..1], 1));
    }

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
