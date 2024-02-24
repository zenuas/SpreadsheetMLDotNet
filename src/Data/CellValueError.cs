namespace SpreadsheetMLDotNet.Data;

public class CellValueError : ICellValue
{
    public static readonly CellValueError DIV_0 = new() { Value = ErrorValues.DIV_0 };
    public static readonly CellValueError GETTING_DATA = new() { Value = ErrorValues.GETTING_DATA };
    public static readonly CellValueError NA = new() { Value = ErrorValues.NA };
    public static readonly CellValueError NAME = new() { Value = ErrorValues.NAME };
    public static readonly CellValueError NULL = new() { Value = ErrorValues.NULL };
    public static readonly CellValueError REF = new() { Value = ErrorValues.REF };
    public static readonly CellValueError VALUE = new() { Value = ErrorValues.VALUE };

    public required ErrorValues Value { get; init; }

    protected CellValueError() { }
}
