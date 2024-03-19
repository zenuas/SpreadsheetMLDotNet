namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class ColorFilter : IFilterColumns
{
    public string DifferentialFormatRecordId { get; set; } = "";
    public bool? FilterByCellColor { get; set; }
}
