using System;

namespace SpreadsheetMLDotNet.Data;

public class CellStyle : IEquatable<CellStyle>
{
    public Font? Font { get; init; }
    public Fill? Fill { get; init; }
    public Border? Border { get; init; }

    public bool Equals(CellStyle? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if ((Font is { } && !Font.Equals(other.Font)) || (Font is null && other.Font is { })) return false;
        if ((Fill is { } && !Fill.Equals(other.Fill)) || (Fill is null && other.Fill is { })) return false;
        if ((Border is { } && !Border.Equals(other.Border)) || (Border is null && other.Border is { })) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as CellStyle);

    public override int GetHashCode() => HashCode.Combine(Font, Fill, Border);
}
