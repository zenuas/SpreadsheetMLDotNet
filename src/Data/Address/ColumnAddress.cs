namespace SpreadsheetMLDotNet.Data.Address;

public readonly struct ColumnAddress : IAddress
{
    public required int Column { get; init; }


    public static implicit operator ColumnAddress(int column) => new() { Column = column };

    public static implicit operator ColumnAddress(string column) => new() { Column = SpreadsheetML.ConvertColumnNameToIndex(column) };
}
