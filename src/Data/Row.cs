using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class Row
{
    public int StartIndex { get; set; } = 0;
    public List<Cell> Values { get; init; } = [];

    public Cell? GetCell(int index) => index < StartIndex || index > StartIndex + Values.Count ? null : Values[index - StartIndex];

    public void SetCell(int index, Cell cell)
    {
        if (index < 1) throw new ArgumentException(nameof(index));

        if (StartIndex < 1 || index < StartIndex)
        {
            StartIndex = index;
            if (StartIndex >= 1 && index + 1 < StartIndex) Values.InsertRange(0, Lists.Repeat(0).Take(StartIndex - index - 1).Select(_ => new Cell { Value = CellValueNull.Instance }));
            Values.Insert(0, cell);
        }
        else if (index > StartIndex + Values.Count - 1)
        {
            if (index > StartIndex + Values.Count - 1) Values.AddRange(Lists.Repeat(0).Take(index - StartIndex - Values.Count).Select(_ => new Cell { Value = CellValueNull.Instance }));
            Values.Add(cell);
        }
        else
        {
            Values[index - StartIndex] = cell;
        }
    }
}
