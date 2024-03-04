using System;

namespace SpreadsheetMLDotNet.Data.Styles;

public class CellStyle : IEquatable<CellStyle>
{
    public Font? Font { get; init; }
    public Fill? Fill { get; init; }
    public Border? Border { get; init; }
    public Alignment? Alignment { get; init; }

    public bool Equals(CellStyle? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(Font, other.Font)) return false;
        if (!Equals(Fill, other.Fill)) return false;
        if (!Equals(Border, other.Border)) return false;
        if (!Equals(Alignment, other.Alignment)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as CellStyle);

    public override int GetHashCode() => HashCode.Combine(Font, Fill, Border, Alignment);
}
