using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.Styles;

public enum UnderlineTypes
{
    [Alias("double")]
    DoubleUnderline,

    [Alias("doubleAccounting")]
    AccountingDoubleUnderline,

    [Alias("none")]
    None,

    [Alias("single")]
    SingleUnderline,

    [Alias("singleAccounting")]
    AccountingSingleUnderline,
}
