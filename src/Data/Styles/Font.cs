using System;

namespace SpreadsheetMLDotNet.Data.Styles;

public class Font : IEquatable<Font>
{
    public bool Equals(Font? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as Font);

    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
}
