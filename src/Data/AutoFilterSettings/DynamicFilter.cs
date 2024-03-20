using System;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class DynamicFilter : IFilterColumns
{
    public DynamicFilterTypes DynamicFilterType { get; init; }
    public double? Value { get; init; }
    public DateTime? ISOValue { get; init; }
    public DateTime? MaxISOValue { get; init; }
}
