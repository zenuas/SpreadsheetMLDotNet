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
        var attr = new Dictionary<string, object?>();

        foreach (var (reader, hierarchy) in XmlReader.Create(worksheet)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "worksheet/cols/col/:START":
                    {
                        var col = new Column();
                        if (reader.GetAttribute("hidden") is { } hidden) col.Hidden = ToBool(hidden);
                        if (reader.GetAttribute("outlineLevel") is { } outlineLevel) col.OutlineLevel = ToUInt(outlineLevel);
                        if (reader.GetAttribute("collapsed") is { } collapsed) col.Collapsed = ToBool(collapsed);
                        if (reader.GetAttribute("phonetic") is { } phonetic) col.ShowPhonetic = ToBool(phonetic);
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
                        if (reader.GetAttribute("hidden") is { } hidden) row.Hidden = ToBool(hidden);
                        if (reader.GetAttribute("outlineLevel") is { } outlineLevel) row.OutlineLevel = ToUInt(outlineLevel);
                        if (reader.GetAttribute("collapsed") is { } collapsed) row.Collapsed = ToBool(collapsed);
                        if (reader.GetAttribute("thickTop") is { } thickTop) row.ThickTop = ToBool(thickTop);
                        if (reader.GetAttribute("thickBot") is { } thickBot) row.ThickBottom = ToBool(thickBot);
                        if (reader.GetAttribute("ph") is { } ph) row.ShowPhonetic = ToBool(ph);
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
                    attr.Clear();
                    break;

                case "worksheet/sheetData/row/c/is/r/:END":
                    cell!.Value.Cast<CellValueInlineString>().Values.Add(new()
                    {
                        Text = ToString(attr.GetOrNull("Text")),
                        FontName = ToString(attr.GetOrNull("FontName")),
                        CharacterSet = ToEnum<CharacterSets>(attr.GetOrNull("CharacterSet")),
                        FontFamily = ToEnum<FontFamilies>(attr.GetOrNull("FontFamily")),
                        Bold = ToBool(attr.GetOrNull("Bold")),
                        Italic = ToBool(attr.GetOrNull("Italic")),
                        StrikeThrough = ToBool(attr.GetOrNull("StrikeThrough")),
                        Outline = ToBool(attr.GetOrNull("Outline")),
                        Shadow = ToBool(attr.GetOrNull("Shadow")),
                        Condense = ToBool(attr.GetOrNull("Condense")),
                        Extend = ToBool(attr.GetOrNull("Extend")),
                        Color = ToColor(attr.GetOrNull("Color")),
                        FontSize = ToDouble(attr.GetOrNull("FontSize")),
                        Underline = ToEnum<UnderlineTypes>(attr.GetOrNull("Underline")),
                        VerticalAlignment = ToEnum<VerticalPositioningLocations>(attr.GetOrNull("VerticalAlignment")),
                        Scheme = ToEnum<FontSchemeStyles>(attr.GetOrNull("Scheme")),
                    });
                    break;

                case "worksheet/sheetData/row/c/is/r/rPr/name/:START": attr["FontName"] = reader.GetAttribute("val")!; break;
                case "worksheet/sheetData/row/c/is/r/rPr/charset/:START": attr["CharacterSet"] = ToEnum<CharacterSets>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/family/:START": attr["FontFamily"] = ToEnum<FontFamilies>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/b/:START": attr["Bold"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/i/:START": attr["Italic"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/strike/:START": attr["StrikeThrough"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/outline/:START": attr["Outline"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/shadow/:START": attr["Shadow"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/condense/:START": attr["Condense"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/extend/:START": attr["Extend"] = ToBool(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/color/:START": attr["Color"] = ToColor(reader.GetAttribute("rgb")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/sz/:START": attr["FontSize"] = ToDouble(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/u/:START": attr["Underline"] = ToEnumAlias<UnderlineTypes>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/vertAlign/:START": attr["VerticalAlignment"] = ToEnumAlias<VerticalPositioningLocations>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/rPr/scheme/:START": attr["Scheme"] = ToEnumAlias<FontSchemeStyles>(reader.GetAttribute("val")!); break;
                case "worksheet/sheetData/row/c/is/r/t/:TEXT": attr["Text"] = reader.Value; break;

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
