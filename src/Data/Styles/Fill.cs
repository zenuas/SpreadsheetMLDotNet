using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class Fill : IEquatable<Fill>
{
    public PatternTypes PatternType { get; set; } = PatternTypes.Solid;
    public Color? ForegroundColor { get; set; }
    public Color? BackgroundColor { get; set; }

    public override bool Equals(object? obj) => Equals(obj as Fill);

    public bool Equals(Fill? other) =>
        other is not null &&
        PatternType == other.PatternType &&
        EqualityComparer<Color?>.Default.Equals(ForegroundColor, other.ForegroundColor) &&
        EqualityComparer<Color?>.Default.Equals(BackgroundColor, other.BackgroundColor);

    public override int GetHashCode() => HashCode.Combine(PatternType, ForegroundColor, BackgroundColor);
}
