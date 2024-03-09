using System.Drawing;

namespace SpreadsheetMLDotNet.Data.Styles;

public struct BorderPropertiesType()
{
    public Color? Color { get; set; }
    public BorderStyles Style { get; set; } = BorderStyles.None;
}