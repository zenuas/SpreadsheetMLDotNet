namespace SpreadsheetMLDotNet.Data.Address;

public readonly struct ColumnAddressRange : IAddressRange
{
    public required int From { get; init; }
    public required int To { get; init; }


    public static implicit operator ColumnAddressRange((int From, int To) range) => new() { From = range.From, To = range.To };

    public static implicit operator ColumnAddressRange((string From, string To) range) => new() { From = SpreadsheetML.ConvertColumnNameToIndex(range.From), To = SpreadsheetML.ConvertColumnNameToIndex(range.To) };

    public static implicit operator ColumnAddressRange(string range) => SpreadsheetML.ConvertColumnRange(range);
}
