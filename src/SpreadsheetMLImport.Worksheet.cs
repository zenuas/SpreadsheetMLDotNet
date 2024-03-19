using Mina.Extension;
using SpreadsheetMLDotNet.Data.SharedStringTable;
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
    public static Worksheet ReadWorksheet(Stream worksheet, string sheet_name, Dictionary<int, IStringItem> shared_strings, CellStyle[] cellstyles)
    {
        var sheet = new Worksheet { Name = sheet_name };
        Row? row = null;
        Cell? cell = null;
        CellTypes cell_type = CellTypes.Number;
        int column = 0;
        RichText? rt = null;

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
                    cell!.Value = new CellValueFormula { Value = reader.Value };
                    break;

                case "worksheet/sheetData/row/c/is/:START":
                    cell = new Cell() { Value = new CellValueInlineString() };
                    cell_type = CellTypes.InlineString;
                    break;

                case "worksheet/sheetData/row/c/is/r/:START":
                    cell!.Value.Cast<CellValueInlineString>().Values.Add(rt = new() { Text = "" });
                    break;

                case "worksheet/sheetData/row/c/is/r/:END":
                    rt = null;
                    break;

                case "worksheet/sheetData/row/c/is/r/rPr/name/:START": rt!.FontName = reader.GetAttribute("val")!; break;
                case "worksheet/sheetData/row/c/is/r/rPr/charset/:START": rt!.CharacterSet = ToEnum<CharacterSets>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/family/:START": rt!.FontFamily = ToEnum<FontFamilies>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/b/:START": rt!.Bold = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/i/:START": rt!.Italic = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/strike/:START": rt!.StrikeThrough = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/outline/:START": rt!.Outline = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/shadow/:START": rt!.Shadow = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/condense/:START": rt!.Condense = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/extend/:START": rt!.Extend = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/color/:START": rt!.Color = ToColor(reader.GetAttribute("rgb")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/sz/:START": rt!.FontSize = ToDouble(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/u/:START": rt!.Underline = ToEnumAlias<UnderlineTypes>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/vertAlign/:START": rt!.VerticalAlignment = ToEnumAlias<VerticalPositioningLocations>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/scheme/:START": rt!.Scheme = ToEnumAlias<FontSchemeStyles>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/t/:TEXT": rt!.Text = reader.Value; break;

                case "worksheet/sheetData/row/c/v/:TEXT":
                    if (cell!.Value is not CellValueFormula) cell!.Value = ToCellValue(reader.Value, cell_type, shared_strings);
                    break;

                case "worksheet/sheetData/row/c/:END":
                    row!.Cells.SetValue(column, cell!);
                    cell = null;
                    column = 0;
                    break;

                case "worksheet/mergeCells/mergeCell/:START":
                    sheet.Merge(SpreadsheetML.ConvertAnyRange(reader.GetAttribute("ref")!));
                    break;
            }
        }
        return sheet;
    }

    public static ICellValue ToCellValue(string value, CellTypes cell_type, Dictionary<int, IStringItem> shared_strings) => cell_type switch
    {
        CellTypes.Boolean => new CellValueBoolean { Value = value == "1" },
        CellTypes.Date => new CellValueDate { Value = DateTime.Parse(value) },
        CellTypes.Error => CellValueError.GetValue(ToEnumAlias<ErrorValues>(value)!.Value),
        CellTypes.String => new CellValueString { Value = value },
        CellTypes.Number => new CellValueDouble { Value = double.Parse(value) },
        CellTypes.SharedString => shared_strings[int.Parse(value)] is RunProperties rpr ? new CellValueInlineString { Values = rpr } : new CellValueString { Value = shared_strings[int.Parse(value)].ToString() },
        _ => throw new ArgumentException(null, nameof(cell_type)),
    };
}
