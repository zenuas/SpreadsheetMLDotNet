namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class DateGroupItem : IFilter
{
    public required uint Year { get; set; }
    public uint Month { get; set; }
    public uint Day { get; set; }
    public uint Hour { get; set; }
    public uint Minute { get; set; }
    public uint Second { get; set; }
    public DateTimeGroupings DateTimeGrouping { get; set; }
}
