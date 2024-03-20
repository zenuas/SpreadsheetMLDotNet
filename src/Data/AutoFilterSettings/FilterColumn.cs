namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class FilterColumn
{
    public uint FilterColumnData { get; init; }
    public bool HiddenAutoFilterButton { get; init; }
    public bool ShowFilterButton { get; init; }
    public IFilterColumns? Filters { get; init; }
}
