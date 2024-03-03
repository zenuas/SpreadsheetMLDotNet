using System;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class Border : IEquatable<Border>
{
    public BorderPropertiesType? Start { get; set; }
    public BorderPropertiesType? End { get; set; }
    public BorderPropertiesType? Top { get; set; }
    public BorderPropertiesType? Bottom { get; set; }
    public BorderPropertiesType? Diagonal { get; set; }
    public BorderPropertiesType? Vertical { get; set; }
    public BorderPropertiesType? Horizontal { get; set; }

    public Border() { }

    public Border(Borders borders, BorderStyles style, Color? color) => SetBorder(borders, style, color);

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

    public bool Equals(Border? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(Start, other.Start)) return false;
        if (!Equals(End, other.End)) return false;
        if (!Equals(Top, other.Top)) return false;
        if (!Equals(Bottom, other.Bottom)) return false;
        if (!Equals(Diagonal, other.Diagonal)) return false;
        if (!Equals(Vertical, other.Vertical)) return false;
        if (!Equals(Horizontal, other.Horizontal)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as Border);

    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
}
