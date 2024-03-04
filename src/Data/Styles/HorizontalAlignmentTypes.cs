using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Styles;

public enum HorizontalAlignmentTypes
{
    [Alias("center")]
    CenteredHorizontalAlignment,

    [Alias("centerContinuous")]
    CenterContinuousHorizontalAlignment,

    [Alias("distributed")]
    DistributedHorizontalAlignment,

    [Alias("fill")]
    Fill,

    [Alias("general")]
    GeneralHorizontalAlignment,

    [Alias("justify")]
    Justify,

    [Alias("left")]
    LeftHorizontalAlignment,

    [Alias("right")]
    RightHorizontalAlignment,
}
