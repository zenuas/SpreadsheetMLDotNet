namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class CustomFilter
{
    public FilterOperators? Operator { get; init; }
    public string Value { get; init; } = "";
}
