namespace SpreadsheetMLDotNet.Data.SharedStringTable;

public class Text : IStringItem
{
    public required string Value { get; init; }

    public override string ToString() => Value;
}
