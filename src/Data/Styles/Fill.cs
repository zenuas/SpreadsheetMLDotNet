using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct Fill
{
    public PatternTypes? PatternType { get; init; }
    public Color? ForegroundColor { get; init; }
    public Color? BackgroundColor { get; init; }
}
