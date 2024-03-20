namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class ColorFilter : IFilterColumns
{
    public string DifferentialFormatRecordId { get; init; } = "";
    public bool? FilterByCellColor { get; init; }
}
