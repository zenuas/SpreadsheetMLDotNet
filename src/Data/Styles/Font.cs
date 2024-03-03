using System;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class Font : IEquatable<Font>
{
    public string FontName { get; set; } = "";
    public CharacterSets? CharacterSet { get; set; }
    public FontFamilies? FontFamily { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public bool? StrikeThrough { get; set; }
    public bool? Outline { get; set; }
    public bool? Shadow { get; set; }
    public bool? Condense { get; set; }
    public bool? Extend { get; set; }
    public Color? Color { get; set; }
    public double? FontSize { get; set; }
    public UnderlineTypes? Underline { get; set; }
    public VerticalPositioningLocations? VerticalAlignment { get; set; }
    public FontSchemeStyles? Scheme { get; set; }

    public bool Equals(Font? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(FontName, other.FontName)) return false;
        if (!Equals(CharacterSet, other.CharacterSet)) return false;
        if (!Equals(FontFamily, other.FontFamily)) return false;
        if (!Equals(Bold, other.Bold)) return false;
        if (!Equals(Italic, other.Italic)) return false;
        if (!Equals(StrikeThrough, other.StrikeThrough)) return false;
        if (!Equals(Outline, other.Outline)) return false;
        if (!Equals(Shadow, other.Shadow)) return false;
        if (!Equals(Condense, other.Condense)) return false;
        if (!Equals(Extend, other.Extend)) return false;
        if (!Equals(Color, other.Color)) return false;
        if (!Equals(FontSize, other.FontSize)) return false;
        if (!Equals(Underline, other.Underline)) return false;
        if (!Equals(VerticalAlignment, other.VerticalAlignment)) return false;
        if (!Equals(Scheme, other.Scheme)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as Font);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(FontName);
        hash.Add(CharacterSet);
        hash.Add(FontFamily);
        hash.Add(Bold);
        hash.Add(Italic);
        hash.Add(StrikeThrough);
        hash.Add(Outline);
        hash.Add(Shadow);
        hash.Add(Condense);
        hash.Add(Extend);
        hash.Add(Color);
        hash.Add(FontSize);
        hash.Add(Underline);
        hash.Add(VerticalAlignment);
        hash.Add(Scheme);
        return hash.ToHashCode();
    }
}
