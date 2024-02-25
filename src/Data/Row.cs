using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class Row
{
    public int StartCellIndex { get; set; } = 0;
    public List<Cell> Values { get; init; } = [];

    public Cell? GetCell(int index) => index < StartCellIndex || index > StartCellIndex + Values.Count ? null : Values[index - StartCellIndex];

    public void SetCell(int index, Cell cell)
    {
        if (index < 1) throw new ArgumentException(nameof(index));

        if (StartCellIndex < 1 || index < StartCellIndex)
        {
            StartCellIndex = index;
            if (StartCellIndex >= 1 && index + 1 < StartCellIndex) Values.InsertRange(0, Lists.Repeat(0).Take(StartCellIndex - index - 1).Select(_ => new Cell { Value = CellValueNull.Instance }));
            Values.Insert(0, cell);
        }
        else if (index > StartCellIndex + Values.Count - 1)
        {
            if (index > StartCellIndex + Values.Count - 1) Values.AddRange(Lists.Repeat(0).Take(index - StartCellIndex - Values.Count).Select(_ => new Cell { Value = CellValueNull.Instance }));
            Values.Add(cell);
        }
        else
        {
            Values[index - StartCellIndex] = cell;
        }
    }
}
