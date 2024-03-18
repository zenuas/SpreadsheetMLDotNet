using System;

namespace SpreadsheetMLDotNet.Data.Styles;

public class NumberFormatId : INumberFormat, IEquatable<NumberFormatId>
{
    public required NumberFormats FormatId { get; set; }

    public override bool Equals(object? obj) => Equals(obj as NumberFormatId);

    public bool Equals(NumberFormatId? other) =>
        other is not null &&
        FormatId == other.FormatId;

    public override int GetHashCode() => HashCode.Combine(FormatId);
}
