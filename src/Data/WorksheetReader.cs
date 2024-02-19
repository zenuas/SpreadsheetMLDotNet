using System;
using System.Collections;
using System.Collections.Generic;

namespace SpreadsheetMLReader.Data;

public class WorksheetReader : IDisposable, IEnumerable<(string Cell, object Value)>, IEnumerable
{
    public required Func<IEnumerator<(string Cell, object Value)>> GetEnumerator_ { get; init; }
    public required Action Dispose { get; init; }


    IEnumerator<(string Cell, object Value)> IEnumerable<(string Cell, object Value)>.GetEnumerator() => GetEnumerator_();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator_();
    void IDisposable.Dispose()
    {
        Dispose();
        GC.SuppressFinalize(this);
    }
}
