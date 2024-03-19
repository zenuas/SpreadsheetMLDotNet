using System;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class DynamicFilter : IFilterColumns
{
    public DynamicFilterTypes DynamicFilterType { get; set; }
    public double? Value { get; set; }
    public DateTime? ISOValue { get; set; }
    public DateTime? MaxISOValue { get; set; }
}
