using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Styles;

public enum BorderStyles
{
    [Alias("none")]
    None,

    [Alias("thin")]
    Thin,

    [Alias("medium")]
    Medium,

    [Alias("dashed")]
    Dashed,

    [Alias("dotted")]
    Dotted,

    [Alias("thick")]
    Thick,

    [Alias("double")]
    Double,

    [Alias("hair")]
    Hair,

    [Alias("mediumDashed")]
    MediumDashed,

    [Alias("dashDot")]
    DashDot,

    [Alias("mediumDashDot")]
    MediumDashDot,

    [Alias("dashDotDot")]
    DashDotDot,

    [Alias("mediumDashDotDot")]
    MediumDashDotDot,

    [Alias("slantDashDot")]
    SlantDashDot,
}
