namespace SpreadsheetMLDotNet.Data.Address;

public readonly struct RowAddress : IAddress
{
    public required int Row { get; init; }


    public static implicit operator RowAddress(int row) => new() { Row = row };
}
