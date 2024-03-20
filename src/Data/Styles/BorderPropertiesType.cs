using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct BorderPropertiesType
{
    public Color? Color { get; init; }
    public BorderStyles? Style { get; init; }
}