using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public class RichText : IFont
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
}
