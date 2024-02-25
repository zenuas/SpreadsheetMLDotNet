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
        if (index < 1) throw new ArgumentException(nameof(index));

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
}
