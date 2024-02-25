﻿using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class Row
{
    public int StartCellIndex { get; set; } = 0;
    public List<Cell> Cells { get; init; } = [];

    public Cell GetCell(int index) => index < StartCellIndex || index > StartCellIndex + Cells.Count - 1
        ? new Cell { Value = CellValueNull.Instance }.Return(x => SetCell(index, x))
        : Cells[index - StartCellIndex];

    public void SetCell(int index, Cell cell)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        if (StartCellIndex < 1 || index < StartCellIndex)
        {
            StartCellIndex = index;
            if (StartCellIndex >= 1 && index + 1 < StartCellIndex) Cells.InsertRange(0, Lists.Repeat(0).Take(StartCellIndex - index - 1).Select(_ => new Cell { Value = CellValueNull.Instance }));
            Cells.Insert(0, cell);
        }
        else if (index > StartCellIndex + Cells.Count - 1)
        {
            if (index > StartCellIndex + Cells.Count - 1) Cells.AddRange(Lists.Repeat(0).Take(index - StartCellIndex - Cells.Count).Select(_ => new Cell { Value = CellValueNull.Instance }));
            Cells.Add(cell);
        }
        else
        {
            Cells[index - StartCellIndex] = cell;
        }
    }
}
