namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class CustomFilter
{
    public FilterOperators? Operator { get; set; }
    public string Value { get; set; } = "";
}
