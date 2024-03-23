using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public enum SheetStates
{
    [Alias("visible")]
    Visible,

    [Alias("hidden")]
    Hidden,

    [Alias("veryHidden")]
    VeryHidden,
}
