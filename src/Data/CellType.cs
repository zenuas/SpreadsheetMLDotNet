using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data;

public enum CellType
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
