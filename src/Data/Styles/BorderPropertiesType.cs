using System;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class BorderPropertiesType : IEquatable<BorderPropertiesType>
{
    public Color? Color { get; set; }
    public BorderStyles Style { get; set; } = BorderStyles.None;

    public bool Equals(BorderPropertiesType? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(Color, other.Color)) return false;
        if (!Equals(Style, other.Style)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as BorderPropertiesType);

    public override int GetHashCode() => HashCode.Combine(Color, Style);
}