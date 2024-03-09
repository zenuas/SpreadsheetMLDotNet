using System;

namespace SpreadsheetMLDotNet.Data.Styles;

public class NumberFormatId : INumberFormat, IEquatable<NumberFormatId>
{
    public required NumberFormats FormatId { get; set; }

    public bool Equals(NumberFormatId? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(FormatId, other.FormatId)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as NumberFormatId);

    public override int GetHashCode() => HashCode.Combine(FormatId);
}
