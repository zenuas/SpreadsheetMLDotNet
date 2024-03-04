using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Styles;

public enum VerticalAlignmentTypes
{
    [Alias("bottom")]
    AlignedToBottom,

    [Alias("center")]
    CenteredVerticalAlignment,

    [Alias("distributed")]
    DistributedVerticalAlignment,

    [Alias("justify")]
    JustifiedVertically,

    [Alias("top")]
    AlignTop,
}
