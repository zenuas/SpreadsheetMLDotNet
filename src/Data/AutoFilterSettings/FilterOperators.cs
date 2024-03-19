using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public enum FilterOperators
{
    [Alias("equal")]
    Equal,

    [Alias("lessThan")]
    LessThan,

    [Alias("lessThanOrEqual")]
    LessThanOrEqual,

    [Alias("notEqual")]
    NotEqual,

    [Alias("greaterThanOrEqual")]
    GreaterThanOrEqual,

    [Alias("greaterThan")]
    reaterThan,
}
