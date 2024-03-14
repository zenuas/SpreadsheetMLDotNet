using Mina.Extension;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Worksheet : IRelationshipable
{
    public required string Name { get; set; }
    public IndexedList<Row> Rows { get; init; } = new() { New = () => new() };
    public IndexedList<Column> Columns { get; init; } = new() { New = () => new() };

    public Row? GetRowOrDefault(int row) => Rows.GetOrDefault(row);

    public Row GetRow(int row) => Rows.GetOrNewAdd(row);

    public Column? GetColumnOrDefault(string column) => Columns.GetOrDefault(SpreadsheetML.ConvertColumnNameToIndex(column));

    public Column? GetColumnOrDefault(int column) => Columns.GetOrDefault(column);

    public Column GetColumn(string column) => Columns.GetOrNewAdd(SpreadsheetML.ConvertColumnNameToIndex(column));

    public Column GetColumn(int column) => Columns.GetOrNewAdd(column);

    public Cell? GetCellOrDefault(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCellOrDefault(x.Row, x.Column));

    public Cell? GetCellOrDefault(int row, int column) => Rows.GetOrDefault(row)?.Cells?.GetOrDefault(column);

    public Cell GetCell(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCell(x.Row, x.Column));

    public Cell GetCell(int row, int column) => Rows.GetOrNewAdd(row).Cells.GetOrNewAdd(column);

    public void SetCell(string address, Cell cell) => SpreadsheetML.ConvertCellAddress(address).Return(x => SetCell(x.Row, x.Column, cell));

    public void SetCell(int row, int column, Cell cell) => Rows.GetOrNewAdd(row).Cells.SetValue(column, cell);
}
