using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public struct Fill()
{
    public PatternTypes PatternType { get; set; } = PatternTypes.Solid;
    public Color? ForegroundColor { get; set; }
    public Color? BackgroundColor { get; set; }
}
