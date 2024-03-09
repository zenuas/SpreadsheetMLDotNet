using System;

namespace SpreadsheetMLDotNet.Data.Styles;

public class NumberFormatCode : INumberFormat, IEquatable<NumberFormatCode>
{
    public required string FormatCode { get; set; }

    public bool Equals(NumberFormatCode? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(FormatCode, other.FormatCode)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as NumberFormatCode);

    public override int GetHashCode() => HashCode.Combine(FormatCode);
}
