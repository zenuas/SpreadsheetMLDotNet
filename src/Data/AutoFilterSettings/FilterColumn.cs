namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class FilterColumn
{
    public uint FilterColumnData { get; set; }
    public bool HiddenAutoFilterButton { get; set; }
    public bool ShowFilterButton { get; set; }
    public IFilterColumns? Filters { get; set; }
}
