using System;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.Styles;

public class CellStyle : IEquatable<CellStyle>
{
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
    public INumberFormat? NumberFormat { get; set; }

    public override bool Equals(object? obj) => Equals(obj as CellStyle);

    public bool Equals(CellStyle? other) =>
        other is not null &&
        EqualityComparer<Font?>.Default.Equals(Font, other.Font) &&
        EqualityComparer<Fill?>.Default.Equals(Fill, other.Fill) &&
        EqualityComparer<Border?>.Default.Equals(Border, other.Border) &&
        EqualityComparer<Alignment?>.Default.Equals(Alignment, other.Alignment) &&
        EqualityComparer<INumberFormat?>.Default.Equals(NumberFormat, other.NumberFormat);

    public override int GetHashCode() => HashCode.Combine(Font, Fill, Border, Alignment, NumberFormat);
}
