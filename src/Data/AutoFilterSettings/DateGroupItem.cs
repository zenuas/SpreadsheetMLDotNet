namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class DateGroupItem : IFilter
{
    public required uint Year { get; init; }
    public uint Month { get; init; }
    public uint Day { get; init; }
    public uint Hour { get; init; }
    public uint Minute { get; init; }
    public uint Second { get; init; }
    public required DateTimeGroupings DateTimeGrouping { get; init; }
}
