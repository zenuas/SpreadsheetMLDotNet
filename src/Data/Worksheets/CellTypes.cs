using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public enum CellTypes
{
    [Alias("b")]
    Boolean,

    [Alias("d")]
    Date,

    [Alias("e")]
    Error,

    [Alias("inlineStr")]
    InlineString,

    [Alias("n")]
    Number,

    [Alias("s")]
    SharedString,

    [Alias("str")]
    String,
}
