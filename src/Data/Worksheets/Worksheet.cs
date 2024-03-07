using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Worksheet : IRelationshipable
{
    public required string Name { get; set; }
    public int StartRowIndex { get; set; } = 0;
    public List<Row> Rows { get; init; } = [];
    public int StartColumnIndex { get; set; } = 0;
    public List<Column> Columns { get; init; } = [];

    public Row? TryGetRow(int index) => index < StartRowIndex || index > StartRowIndex + Rows.Count - 1
        ? null
        : Rows[index - StartRowIndex];

    public Row GetRow(int index) => TryGetRow(index) is { } row
        ? row
        : new Row().Return(x => SetRow(index, x));

    public void SetRow(int index, Row row)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        if (StartRowIndex < 1 || index < StartRowIndex)
        {
            StartRowIndex = index;
            if (StartRowIndex >= 1 && index + 1 < StartRowIndex) Rows.InsertRange(0, Lists.Repeat(0).Take(StartRowIndex - index - 1).Select(_ => new Row()));
            Rows.Insert(0, row);
        }
        else if (index > StartRowIndex + Rows.Count - 1)
        {
            if (index > StartRowIndex + Rows.Count - 1) Rows.AddRange(Lists.Repeat(0).Take(index - StartRowIndex - Rows.Count).Select(_ => new Row()));
            Rows.Add(row);
        }
        else
        {
            Rows[index - StartRowIndex] = row;
        }
    }

    public Column? TryGetColumn(string col) => TryGetColumn(SpreadsheetML.ConvertColumnNameToIndex(col));

    public Column? TryGetColumn(int index) => index < StartColumnIndex || index > StartColumnIndex + Columns.Count - 1
        ? null
        : Columns[index - StartColumnIndex];

    public Column GetColumn(string col) => GetColumn(SpreadsheetML.ConvertColumnNameToIndex(col));

    public Column GetColumn(int index) => TryGetColumn(index) is { } row
        ? row
        : new Column().Return(x => SetColumn(index, x));

    public void SetColumn(int index, Column column)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        if (StartColumnIndex < 1 || index < StartColumnIndex)
        {
            StartColumnIndex = index;
            if (StartColumnIndex >= 1 && index + 1 < StartColumnIndex) Columns.InsertRange(0, Lists.Repeat(0).Take(StartColumnIndex - index - 1).Select(_ => new Column()));
            Columns.Insert(0, column);
        }
        else if (index > StartColumnIndex + Columns.Count - 1)
        {
            if (index > StartColumnIndex + Columns.Count - 1) Columns.AddRange(Lists.Repeat(0).Take(index - StartColumnIndex - Columns.Count).Select(_ => new Column()));
            Columns.Add(column);
        }
        else
        {
            Columns[index - StartColumnIndex] = column;
        }
    }

    public Cell? TryGetCell(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => TryGetCell(x.Row, x.Column));

    public Cell? TryGetCell(int row, int column) => TryGetRow(row)?.TryGetCell(column);

    public Cell GetCell(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCell(x.Row, x.Column));

    public Cell GetCell(int row, int column) => GetRow(row).GetCell(column);

    public void SetCell(string address, Cell cell) => SpreadsheetML.ConvertCellAddress(address).Return(x => SetCell(x.Row, x.Column, cell));

    public void SetCell(int row, int column, Cell cell) => GetRow(row).SetCell(column, cell);
}
