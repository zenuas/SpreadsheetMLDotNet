using Mina.Extension;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Worksheets;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLImport
{
    public static Worksheet ReadWorksheet(Stream worksheet, string sheet_name, Dictionary<int, string> shared_strings, CellStyle[] cellstyles)
    {
        var sheet = new Worksheet { Name = sheet_name };
        Row? row = null;
        Cell? cell = null;
        CellTypes cell_type = CellTypes.Number;
        int column = 0;
        foreach (var (reader, hierarchy) in XmlReader.Create(worksheet)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "worksheet/cols/col/:START":
                    {
                        var col = new Column();
                        if (reader.GetAttribute("customWidth") is { } customWidth && ToBool(customWidth))
                        {
                            if (reader.GetAttribute("width") is { } width) col.Width = ToDouble(width);
                            if (reader.GetAttribute("bestFit") is { } bestFit) col.BestFitColumnWidth = ToBool(bestFit);
                        }
                        if (reader.GetAttribute("style") is { } style) SetStyle(col, cellstyles[ToInt(style)]);
                        sheet.Columns.SetValue(ToInt(reader.GetAttribute("min")!), col);
                        break;
                    }

                case "worksheet/sheetData/row/:START":
                    {
                        row = new Row();
                        if (reader.GetAttribute("customHeight") is { } customHeight && ToBool(customHeight)) row.Height = ToDouble(reader.GetAttribute("ht")!);
                        if (reader.GetAttribute("customFormat") is { } customFormat && ToBool(customFormat)) SetStyle(row, cellstyles[ToInt(reader.GetAttribute("s")!)]);
                        sheet.Rows.SetValue(ToInt(reader.GetAttribute("r")!), row);
                        break;
                    }

                case "worksheet/sheetData/row/:END":
                    row = null;
                    break;

                case "worksheet/sheetData/row/c/:START":
                    {
                        cell = new Cell() { Value = CellValueNull.Instance };
                        column = SpreadsheetML.ConvertCellAddress(reader.GetAttribute("r")!).Column;
                        cell_type = reader.GetAttribute("t") is { } t ? ToEnumAlias<CellTypes>(t)!.Value : CellTypes.Number;
                        if (reader.GetAttribute("s") is { } s) SetStyle(cell, cellstyles[ToInt(s)]);
                        break;
                    }

                case "worksheet/sheetData/row/c/f/:TEXT":
                    {
                        cell!.Value = new CellValueFormula { Value = reader.Value };
                        break;
                    }

                case "worksheet/sheetData/row/c/v/:TEXT":
                    if (cell!.Value is not CellValueFormula) cell!.Value = ToCellValue(reader.Value, cell_type, shared_strings);
                    break;

                case "worksheet/sheetData/row/c/:END":
                    row!.Cells.SetValue(column, cell!);
                    cell = null;
                    column = 0;
                    break;
            }
        }
        return sheet;
    }

    public static ICellValue ToCellValue(string value, CellTypes cell_type, Dictionary<int, string> shared_strings) => cell_type switch
    {
        CellTypes.Boolean => new CellValueBoolean { Value = value == "1" },
        CellTypes.Date => new CellValueDate { Value = DateTime.Parse(value) },
        CellTypes.Error => CellValueError.GetValue(ToEnumAlias<ErrorValues>(value)!.Value),
        CellTypes.InlineString => new CellValueString { Value = value },
        CellTypes.String => new CellValueString { Value = value },
        CellTypes.Number => new CellValueDouble { Value = double.Parse(value) },
        CellTypes.SharedString => new CellValueString { Value = shared_strings[int.Parse(value)] },
        _ => throw new ArgumentException(null, nameof(cell_type)),
    };
}
