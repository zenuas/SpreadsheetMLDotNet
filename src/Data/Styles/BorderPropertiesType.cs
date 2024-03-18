using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class BorderPropertiesType : IEquatable<BorderPropertiesType>
{
    public Color? Color { get; set; }
    public BorderStyles Style { get; set; } = BorderStyles.None;

    public override bool Equals(object? obj) => Equals(obj as BorderPropertiesType);

    public bool Equals(BorderPropertiesType? other) =>
        other is not null &&
        EqualityComparer<Color?>.Default.Equals(Color, other.Color) &&
        Style == other.Style;

    public override int GetHashCode() => HashCode.Combine(Color, Style);
}