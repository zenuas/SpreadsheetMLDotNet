using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Styles;

public enum VerticalPositioningLocations
{
    [Alias("baseline")]
    RegularVerticalPositioning,

    [Alias("subscript")]
    Subscript,

    [Alias("superscript")]
    Superscript,
}
