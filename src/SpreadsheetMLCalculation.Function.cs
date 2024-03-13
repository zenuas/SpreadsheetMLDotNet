using SpreadsheetMLDotNet.Calculation;
using SpreadsheetMLDotNet.Data.Worksheets;
using System;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLCalculation
{
    public static ICellValue EvaluateFunction(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, FunctionCall call)
    {
        switch (call.Name)
        {
            case "SUM": return Sum(calc, current_sheet, call.Arguments);
        }
        return CellValueError.NAME;
    }

    public static ICellValue Sum(Dictionary<string, WorksheetCalculation> calc, Worksheet current_sheet, Span<IFormula> values) =>
        values.Length == 0 ? new CellValueDouble { Value = 0 }
        : Evaluate(values[0], calc, current_sheet) is not CellValueDouble d
            ? CellValueError.VALUE
            : EvaluateDouble(d, "+", Sum(calc, current_sheet, values[1..]));
}
