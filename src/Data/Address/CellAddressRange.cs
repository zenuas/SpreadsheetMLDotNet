namespace SpreadsheetMLDotNet.Data.Address;

public readonly struct CellAddressRange : IAddressRange
{
    public required CellAddress From { get; init; }
    public required CellAddress To { get; init; }


    public static implicit operator CellAddressRange(((int Row, int Column) From, (int Row, int Column) To) range) => new() { From = range.From, To = range.To };

    public static implicit operator CellAddressRange((string From, string To) range) => new() { From = range.From, To = range.To };

    public static implicit operator CellAddressRange(string range) => SpreadsheetML.ConvertCellRange(range);

    public string ToAddressName() => $"{From}:{To}";
}
