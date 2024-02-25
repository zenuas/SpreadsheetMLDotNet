using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class Worksheet : IRelationshipable
{
    public required string Name { get; set; }
    public int StartRowIndex { get; set; } = 0;
    public List<Row> Values { get; init; } = [];

    public Row? GetRow(int index) => index < StartRowIndex || index > StartRowIndex + Values.Count ? null : Values[index - StartRowIndex];

    public void SetRow(int index, Row row)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        if (StartRowIndex < 1 || index < StartRowIndex)
        {
            StartRowIndex = index;
            if (StartRowIndex >= 1 && index + 1 < StartRowIndex) Values.InsertRange(0, Lists.Repeat(0).Take(StartRowIndex - index - 1).Select(_ => new Row()));
            Values.Insert(0, row);
        }
        else if (index > StartRowIndex + Values.Count - 1)
        {
            if (index > StartRowIndex + Values.Count - 1) Values.AddRange(Lists.Repeat(0).Take(index - StartRowIndex - Values.Count).Select(_ => new Row()));
            Values.Add(row);
        }
        else
        {
            Values[index - StartRowIndex] = row;
        }
    }

    public Cell? GetCell(string address) => SpreadsheetML.ConvertCellAddress(address).To(x => GetCell(x.Row, x.Column));

    public Cell? GetCell(int row, int column) => GetRow(row)?.GetCell(column);

    public void SetCell(string address, Cell cell) => SpreadsheetML.ConvertCellAddress(address).Return(x => SetCell(x.Row, x.Column, cell));

    public void SetCell(int row, int column, Cell cell) => (GetRow(row) ?? new Row().Return(x => SetRow(row, x))).SetCell(column, cell);
}
