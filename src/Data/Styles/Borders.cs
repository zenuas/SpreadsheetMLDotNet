using System;

namespace SpreadsheetMLDotNet.Data.Styles;

[Flags]
public enum Borders
{
    Start = 1,
    End = 1 << 1,
    Top = 1 << 2,
    Bottom = 1 << 3,
    Diagonal = 1 << 4,
    Vertical = 1 << 5,
    Horizontal = 1 << 6,
}
