using Mina.Extension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data;

public class IndexedList<T> : List<T> where T : class
{
    public required Func<T> New { get; init; }
    public int StartIndex { get; set; } = 0;

    public T? GetOrDefault(int index) => index < StartIndex || index > StartIndex + Count - 1
        ? null
        : this[index - StartIndex];

    public T GetOrNewAdd(int index) => GetOrDefault(index) is { } value
        ? value
        : New().Return(x => SetValue(index, x));

    public void SetValue(int index, T value)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        if (StartIndex < 1 || index < StartIndex)
        {
            if (StartIndex >= 1 && index + 1 < StartIndex) InsertRange(0, Lists.Repeat(0).Take(StartIndex - index - 1).Select(_ => New()));
            StartIndex = index;
            Insert(0, value);
        }
        else if (index > StartIndex + Count - 1)
        {
            if (index > StartIndex + Count) AddRange(Lists.Repeat(0).Take(index - StartIndex - Count).Select(_ => New()));
            Add(value);
        }
        else
        {
            this[index - StartIndex] = value;
        }
    }

    public IEnumerable<(int Index, T Value)> GetEnumerableWithIndex()
    {
        for (var i = 0; i < Count; i++)
        {
            yield return (i + StartIndex, this[i]);
        }
    }
}
