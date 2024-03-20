using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct Font() : IFont
{
    public string FontName { get; init; } = "";
    public CharacterSets? CharacterSet { get; init; }
    public FontFamilies? FontFamily { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? StrikeThrough { get; init; }
    public bool? Outline { get; init; }
    public bool? Shadow { get; init; }
    public bool? Condense { get; init; }
    public bool? Extend { get; init; }
    public Color? Color { get; init; }
    public double? FontSize { get; init; }
    public UnderlineTypes? Underline { get; init; }
    public VerticalPositioningLocations? VerticalAlignment { get; init; }
    public FontSchemeStyles? Scheme { get; init; }
}
