using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct Border()
{
    public BorderPropertiesType? Start { get; init; }
    public BorderPropertiesType? End { get; init; }
    public BorderPropertiesType? Top { get; init; }
    public BorderPropertiesType? Bottom { get; init; }
    public BorderPropertiesType? Diagonal { get; init; }
    public BorderPropertiesType? Vertical { get; init; }
    public BorderPropertiesType? Horizontal { get; init; }

    public Border(Borders borders, BorderStyles style, Color? color) : this()
    {
        if (borders.HasFlag(Borders.Start)) Start = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.End)) End = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Top)) Top = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Bottom)) Bottom = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Diagonal)) Diagonal = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Vertical)) Vertical = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Horizontal)) Horizontal = new() { Style = style, Color = color };
    }
}
