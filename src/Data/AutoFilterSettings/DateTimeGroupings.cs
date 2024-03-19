using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public enum DateTimeGroupings
{
    [Alias("year")]
    Year,

    [Alias("month")]
    Month,

    [Alias("day")]
    Day,

    [Alias("hour")]
    Hour,

    [Alias("minute")]
    Minute,

    [Alias("second")]
    Second,
}
