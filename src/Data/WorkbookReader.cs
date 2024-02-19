using System;

namespace SpreadsheetMLReader.Data;

public class WorkbookReader : IDisposable
{
    public required string[] WorkSheetNames { get; init; }
    public required Func<string, WorksheetReader> OpenWorksheet { get; init; }
    public required Action Dispose { get; init; }

    void IDisposable.Dispose()
    {
        Dispose();
        GC.SuppressFinalize(this);
    }
}
