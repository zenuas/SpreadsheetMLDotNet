using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public struct Border()
{
    public BorderPropertiesType? Start { get; set; }
    public BorderPropertiesType? End { get; set; }
    public BorderPropertiesType? Top { get; set; }
    public BorderPropertiesType? Bottom { get; set; }
    public BorderPropertiesType? Diagonal { get; set; }
    public BorderPropertiesType? Vertical { get; set; }
    public BorderPropertiesType? Horizontal { get; set; }

    public Border(Borders borders, BorderStyles style, Color? color) : this() => SetBorder(borders, style, color);

    public void SetBorder(Borders borders, BorderStyles style, Color? color)
    {
        if (borders.HasFlag(Borders.Start)) Start = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.End)) End = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Top)) Top = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Bottom)) Bottom = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Diagonal)) Diagonal = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Vertical)) Vertical = new() { Style = style, Color = color };
        if (borders.HasFlag(Borders.Horizontal)) Horizontal = new() { Style = style, Color = color };
    }

    public void Clear()
    {
        Start = null;
        End = null;
        Top = null;
        Bottom = null;
        Diagonal = null;
        Vertical = null;
        Horizontal = null;
    }
}
