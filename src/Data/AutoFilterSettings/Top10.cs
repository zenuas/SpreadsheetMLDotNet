namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class Top10
{
    public double? FilterValue { get; set; }
    public bool? FilterByPercent { get; set; }
    public bool? Top { get; set; }
    public required double Value { get; set; }
}
