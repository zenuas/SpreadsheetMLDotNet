﻿using SpreadsheetMLDotNet.Calculation;
using Xunit;

namespace SpreadsheetMLDotNet.Test;

public class SpreadsheetMLParserTest
{
    [Fact]
    public void ParseTokensTest()
    {
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens(""), []);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("0"), [(TokenTypes.Number, "0")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("0123"), [(TokenTypes.Number, "0123")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("a"), [(TokenTypes.Token, "a")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("abc123"), [(TokenTypes.Token, "abc123")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("\"123\""), [(TokenTypes.String, "123")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("\"123 abc\""), [(TokenTypes.String, "123 abc")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("\"123\"\" abc\""), [(TokenTypes.String, "123\" abc")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("a+1"), [(TokenTypes.Token, "a"), (TokenTypes.Operator, "+"), (TokenTypes.Number, "1")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("a + 1"), [(TokenTypes.Token, "a"), (TokenTypes.Operator, "+"), (TokenTypes.Number, "1")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("a<=1>=b"), [
            (TokenTypes.Token, "a"),
            (TokenTypes.Operator, "<="),
            (TokenTypes.Number, "1"),
            (TokenTypes.Operator, ">="),
            (TokenTypes.Token, "b"),
        ]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("(abc-123+\"xyz\")"), [
            (TokenTypes.LeftParenthesis, "("),
            (TokenTypes.Token, "abc"),
            (TokenTypes.Operator, "-"),
            (TokenTypes.Number, "123"),
            (TokenTypes.Operator, "+"),
            (TokenTypes.String, "xyz"),
            (TokenTypes.RightParenthesis, ")"),
        ]);
    }

    [Fact]
    public void ParseTokensBadTest()
    {
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("\""), [(TokenTypes.String, "")]);
        Assert.Equal<(TokenTypes, string)>(SpreadsheetMLParser.ParseTokens("\"abc"), [(TokenTypes.String, "abc")]);
    }

    [Fact]
    public void ParseTest()
    {
        Assert.Equivalent(SpreadsheetMLParser.Parse("a"), new Token { Value = "a" });
        Assert.Equivalent(SpreadsheetMLParser.Parse("a+1"), new Expression
        {
            Operator = "+",
            Left = new Token { Value = "a" },
            Right = new Number { Value = 1 },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("a+b*c"), new Expression
        {
            Operator = "+",
            Left = new Token { Value = "a" },
            Right = new Expression { Operator = "*", Left = new Token { Value = "b" }, Right = new Token { Value = "c" } },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("a*b+c"), new Expression
        {
            Operator = "+",
            Left = new Expression { Operator = "*", Left = new Token { Value = "a" }, Right = new Token { Value = "b" } },
            Right = new Token { Value = "c" },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("(a+b)*(c-d)"), new Expression
        {
            Operator = "*",
            Left = new Unary { Operator = "()", Value = new Expression { Operator = "+", Left = new Token { Value = "a" }, Right = new Token { Value = "b" } } },
            Right = new Unary { Operator = "()", Value = new Expression { Operator = "-", Left = new Token { Value = "c" }, Right = new Token { Value = "d" } } },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("+a"), new Unary { Operator = "+", Value = new Token { Value = "a" } });
        Assert.Equivalent(SpreadsheetMLParser.Parse("-a*b"), new Expression
        {
            Operator = "*",
            Left = new Unary { Operator = "-", Value = new Token { Value = "a" } },
            Right = new Token { Value = "b" },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("-a*b+c"), new Expression
        {
            Operator = "+",
            Left = new Expression
            {
                Operator = "*",
                Left = new Unary { Operator = "-", Value = new Token { Value = "a" } },
                Right = new Token { Value = "b" },
            },
            Right = new Token { Value = "c" },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("-a+b*c"), new Expression
        {
            Operator = "+",
            Left = new Unary { Operator = "-", Value = new Token { Value = "a" } },
            Right = new Expression
            {
                Operator = "*",
                Left = new Token { Value = "b" },
                Right = new Token { Value = "c" },
            },
        });
    }


    [Fact]
    public void ParseFunctionTest()
    {
        Assert.Equivalent(SpreadsheetMLParser.Parse("SUM()"), new FunctionCall { Name = "SUM", Arguments = [] }, true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("SUM(1)"), new FunctionCall { Name = "SUM", Arguments = [new Number { Value = 1 }] }, true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("SUM(1,2)"), new FunctionCall
        {
            Name = "SUM",
            Arguments = [
                new Number { Value = 1 },
                new Number { Value = 2 },
            ],
        }, true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("SUM(1+2,3)"), new FunctionCall
        {
            Name = "SUM",
            Arguments = [
                new Expression { Operator = "+", Left = new Number { Value = 1 }, Right = new Number { Value = 2 } },
                new Number { Value = 3 },
            ],
        }, true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("(SUM())"), new Unary { Operator = "()", Value = new FunctionCall { Name = "SUM", Arguments = [] } }, true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("-SUM()"), new Unary
        {
            Operator = "-",
            Value = new FunctionCall { Name = "SUM", Arguments = [] },
        }, true);
    }

    [Fact]
    public void ParseBadTest()
    {
        Assert.Equivalent(SpreadsheetMLParser.Parse(")"), new Error());
        Assert.Equivalent(SpreadsheetMLParser.Parse(","), new Error());
        Assert.Equivalent(SpreadsheetMLParser.Parse(":"), new Error());
        Assert.Equivalent(SpreadsheetMLParser.Parse(")a"), new Error());
        Assert.Equivalent(SpreadsheetMLParser.Parse(",a"), new Error());
        Assert.Equivalent(SpreadsheetMLParser.Parse(":a"), new Error());
        Assert.Equivalent(SpreadsheetMLParser.Parse("a)"), new Token { Value = "a" });
        Assert.Equivalent(SpreadsheetMLParser.Parse("a+1)"), new Expression
        {
            Operator = "+",
            Left = new Token { Value = "a" },
            Right = new Number { Value = 1 },
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("(a+1))"), new Unary
        {
            Operator = "()",
            Value = new Expression
            {
                Operator = "+",
                Left = new Token { Value = "a" },
                Right = new Number { Value = 1 },
            }
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("+"), new Unary
        {
            Operator = "+",
            Value = new Null(),
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("a+ +"), new Expression
        {
            Operator = "+",
            Left = new Token { Value = "a" },
            Right = new Unary
            {
                Operator = "+",
                Value = new Null(),
            }
        });
        Assert.Equivalent(SpreadsheetMLParser.Parse("1()"), new Error(), true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("(a)SUM()"), new Error(), true);
        Assert.Equivalent(SpreadsheetMLParser.Parse("(a)()"), new Error(), true);
    }
}
