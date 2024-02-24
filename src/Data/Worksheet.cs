using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class Worksheet : IRelationshipable
{
    public required string Name { get; set; }
    public int StartIndex { get; set; } = 0;
    public List<Row> Values { get; init; } = [];

    public Row? GetRow(int index) => index < StartIndex || index > StartIndex + Values.Count ? null : Values[index - StartIndex];

    public void SetRow(int index, Row row)
    {
        if (index < 1) throw new ArgumentException(nameof(index));

        if (StartIndex < 1 || index < StartIndex)
        {
            StartIndex = index;
            if (StartIndex >= 1 && index + 1 < StartIndex) Values.InsertRange(0, Lists.Repeat(0).Take(StartIndex - index - 1).Select(_ => new Row()));
            Values.Insert(0, row);
        }
        else if (index > StartIndex + Values.Count - 1)
        {
            if (index > StartIndex + Values.Count - 1) Values.AddRange(Lists.Repeat(0).Take(index - StartIndex - Values.Count).Select(_ => new Row()));
            Values.Add(row);
        }
        else
        {
            Values[index - StartIndex] = row;
        }
    }
}
