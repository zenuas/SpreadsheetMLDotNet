using Mina.Attributes;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public enum DynamicFilterTypes
{
    [Alias("null")]
    Null,

    [Alias("aboveAverage")]
    AboveAverage,

    [Alias("belowAverage")]
    BelowAverage,

    [Alias("tomorrow")]
    Tomorrow,

    [Alias("today")]
    Today,

    [Alias("yesterday")]
    Yesterday,

    [Alias("nextWeek")]
    NextWeek,

    [Alias("thisWeek")]
    ThisWeek,

    [Alias("lastWeek")]
    LastWeek,

    [Alias("nextMonth")]
    NextMonth,

    [Alias("thisMonth")]
    ThisMonth,

    [Alias("lastMonth")]
    LastMonth,

    [Alias("nextQuarter")]
    NextQuarter,
    [Alias("thisQuarter")]
    ThisQuarter,

    [Alias("lastQuarter")]
    LastQuarter,

    [Alias("nextYear")]
    NextYear,

    [Alias("thisYear")]
    ThisYear,

    [Alias("lastYear")]
    LastYear,

    [Alias("yearToDate")]
    YearToDate,

    [Alias("Q1")]
    Q1,

    [Alias("Q2")]
    Q2,

    [Alias("Q3")]
    Q3,

    [Alias("Q4")]
    Q4,

    [Alias("M1")]
    M1,

    [Alias("M2")]
    M2,

    [Alias("M3")]
    M3,

    [Alias("M4")]
    M4,

    [Alias("M5")]
    M5,

    [Alias("M6")]
    M6,

    [Alias("M7")]
    M7,

    [Alias("M8")]
    M8,

    [Alias("M9")]
    M9,

    [Alias("M10")]
    M10,

    [Alias("M11")]
    M11,

    [Alias("M12")]
    M12,
}
