using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class RichText : IFont, IEquatable<RichText>
{
    public required string Text { get; set; }
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

    public override bool Equals(object? obj) => Equals(obj as RichText);

    public bool Equals(RichText? other) =>
        other is not null &&
        Text == other.Text &&
        FontName == other.FontName &&
        CharacterSet == other.CharacterSet &&
        FontFamily == other.FontFamily &&
        Bold == other.Bold &&
        Italic == other.Italic &&
        StrikeThrough == other.StrikeThrough &&
        Outline == other.Outline &&
        Shadow == other.Shadow &&
        Condense == other.Condense &&
        Extend == other.Extend &&
        EqualityComparer<Color?>.Default.Equals(Color, other.Color) &&
        FontSize == other.FontSize &&
        Underline == other.Underline &&
        VerticalAlignment == other.VerticalAlignment &&
        Scheme == other.Scheme;

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Text);
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
