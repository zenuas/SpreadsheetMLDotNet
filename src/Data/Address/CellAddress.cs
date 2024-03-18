namespace SpreadsheetMLDotNet.Data.Address;

public readonly struct CellAddress : IAddress
{
    public required int Row { get; init; }
    public required int Column { get; init; }


    public static implicit operator CellAddress((int Row, int Column) cell) => new() { Row = cell.Row, Column = cell.Column };

    public static implicit operator CellAddress(string cell) => SpreadsheetML.ConvertCellAddress(cell);

    public override string ToString() => $"{SpreadsheetML.ConvertColumnIndexToName(Column)}{Row}";
}
