using SpreadsheetMLDotNet.Data.Styles;
using System;

namespace SpreadsheetMLDotNet.Data.Worksheets;

public class Cell : IHaveStyle
{
    public required ICellValue Value { get; set; }
    public Font? Font { get; set; }
    public Fill? Fill { get; set; }
    public Border? Border { get; set; }
    public Alignment? Alignment { get; set; }
    public INumberFormat? NumberFormat { get; set; }

    public static implicit operator Cell(byte x) => new() { Value = new CellValueDouble { Value = x } };

    public static implicit operator Cell(short x) => new() { Value = new CellValueDouble { Value = x } };

    public static implicit operator Cell(int x) => new() { Value = new CellValueDouble { Value = x } };

    public static implicit operator Cell(long x) => new() { Value = new CellValueDouble { Value = x } };

    public static implicit operator Cell(float x) => new() { Value = new CellValueDouble { Value = x } };

    public static implicit operator Cell(double x) => new() { Value = new CellValueDouble { Value = x } };

    public static implicit operator Cell(DateTime x) => new() { Value = new CellValueDate { Value = x } };

    public static implicit operator Cell(bool x) => new() { Value = new CellValueBoolean { Value = x } };

    public static implicit operator Cell(ErrorValues x) => new() { Value = CellValueError.GetValue(x) };

    public static implicit operator Cell(string x) => new() { Value = new CellValueString { Value = x } };
}
