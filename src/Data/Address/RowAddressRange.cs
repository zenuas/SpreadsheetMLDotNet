namespace SpreadsheetMLDotNet.Data.Address;

public readonly struct RowAddressRange : IAddressRange
{
    public required int From { get; init; }
    public required int To { get; init; }


    public static implicit operator RowAddressRange((int From, int To) range) => new() { From = range.From, To = range.To };

    public static implicit operator RowAddressRange(string range) => SpreadsheetML.ConvertRowRange(range);
}
