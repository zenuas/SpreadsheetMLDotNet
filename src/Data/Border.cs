using System;

namespace SpreadsheetMLDotNet.Data;

public class Border : IEquatable<Border>
{
    public bool Equals(Border? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as Border);

    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
}
