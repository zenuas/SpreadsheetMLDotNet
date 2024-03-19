namespace SpreadsheetMLDotNet.Data.SharedStringTable;

public class Text : IStringItem
{
    public required string Value { get; set; }

    public override string ToString() => Value;
}
