namespace SpreadsheetMLDotNet.Data;

public class CellValueNull : ICellValue
{
    public static readonly CellValueNull Instance = new();

    protected CellValueNull() { }
}
