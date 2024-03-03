using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Styles;

public enum PatternTypes
{
    [Alias("none")]
    None,

    [Alias("solid")]
    Solid,

    [Alias("mediumGray")]
    MediumGray,

    [Alias("darkGray")]
    DarkGray,

    [Alias("lightGray")]
    LightGray,

    [Alias("darkHorizontal")]
    DarkHorizontal,

    [Alias("darkVertical")]
    DarkVertical,

    [Alias("darkDown")]
    DarkDown,

    [Alias("darkUp")]
    DarkUp,

    [Alias("darkGrid")]
    DarkGrid,

    [Alias("darkTrellis")]
    DarkTrellis,

    [Alias("lightHorizontal")]
    LightHorizontal,

    [Alias("lightVertical")]
    LightVertical,

    [Alias("lightDown")]
    LightDown,

    [Alias("lightUp")]
    LightUp,

    [Alias("lightGrid")]
    LightGrid,

    [Alias("lightTrellis")]
    LightTrellis,

    [Alias("gray125")]
    Gray125,

    [Alias("gray0625")]
    Gray0625,
}
