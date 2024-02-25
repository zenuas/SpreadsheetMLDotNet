using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class Worksheet : IRelationshipable
{
    public required string Name { get; set; }
    public int StartRowIndex { get; set; } = 0;
    public List<Row> Rows { get; init; } = [];

    public Row? GetRow(int index) => index < StartRowIndex || index > StartRowIndex + Rows.Count - 1 ? null : Rows[index - StartRowIndex];

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

    public Cell? GetCell(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCell(x.Row, x.Column));

    public Cell? GetCell(int row, int column) => GetRow(row)?.GetCell(column);

    public void SetCell(string address, Cell cell) => SpreadsheetML.ConvertCellAddress(address).Return(x => SetCell(x.Row, x.Column, cell));

    public void SetCell(int row, int column, Cell cell) => (GetRow(row) ?? new Row().Return(x => SetRow(row, x))).SetCell(column, cell);
}
