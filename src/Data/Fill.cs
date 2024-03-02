using System;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data;

public class Fill : IEquatable<Fill>
{
    public Color? ForegroundColor { get; set; }
    public Color? BackgroundColor { get; set; }

    public bool Equals(Fill? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (ForegroundColor != other.ForegroundColor) return false;
        if (BackgroundColor != other.BackgroundColor) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as Fill);

    public override int GetHashCode() => HashCode.Combine(ForegroundColor, BackgroundColor);
}
