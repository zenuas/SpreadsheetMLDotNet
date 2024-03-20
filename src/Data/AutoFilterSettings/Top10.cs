namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class Top10
{
    public double? FilterValue { get; init; }
    public bool? FilterByPercent { get; init; }
    public bool? Top { get; init; }
    public required double Value { get; init; }
}
