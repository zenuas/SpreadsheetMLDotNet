using Mina.Extension;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Extension;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLImport
{
    public static CellStyle[] ReadStyles(Stream styles)
    {
        var cellstyles = new List<CellStyle>();
        var formats = new Dictionary<int, NumberFormatCode>();
        var fonts = new List<Font>();
        var fills = new List<Fill>();
        var borders = new List<Border>();
        Font? font = null;
        Fill? fill = null;
        Border? border = null;
        BorderPropertiesType? borderpr = null;
        CellStyle? cellstyle = null;
        bool hasAlignment = false;

        foreach (var (reader, hierarchy) in XmlReader.Create(styles)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "styleSheet/numFmts/numFmt/:START":
                    formats.Add(ToInt(reader.GetAttribute("numFmtId")!), new() { FormatCode = reader.GetAttribute("formatCode")! });
                    break;

                case "styleSheet/fonts/font/:START":
                    font = new();
                    break;

                case "styleSheet/fonts/font/:END":
                    fonts.Add(font!);
                    break;

                case "styleSheet/fonts/font/name/:START": font!.FontName = reader.GetAttribute("val")!; break;
                case "styleSheet/fonts/font/charset/:START": font!.CharacterSet = ToEnum<CharacterSets>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/family/:START": font!.FontFamily = ToEnum<FontFamilies>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/b/:START": font!.Bold = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/i/:START": font!.Italic = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/strike/:START": font!.StrikeThrough = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/outline/:START": font!.Outline = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/shadow/:START": font!.Shadow = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/condense/:START": font!.Condense = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/extend/:START": font!.Extend = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/color/:START": font!.Color = ToColor(reader.GetAttribute("rgb")!); break;
                case "styleSheet/fonts/font/sz/:START": font!.FontSize = ToDouble(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/u/:START": font!.Underline = ToEnumAlias<UnderlineTypes>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/vertAlign/:START": font!.VerticalAlignment = ToEnumAlias<VerticalPositioningLocations>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/scheme/:START": font!.Scheme = ToEnumAlias<FontSchemeStyles>(reader.GetAttribute("val")!); break;

                case "styleSheet/fills/fill/:START":
                    fill = new();
                    break;

                case "styleSheet/fills/fill/:END":
                    fills.Add(fill!);
                    break;

                case "styleSheet/fills/fill/patternFill/:START": fill!.PatternType = ToEnumAlias<PatternTypes>(reader.GetAttribute("patternType")!)!.Value; break;
                case "styleSheet/fills/fill/patternFill/fgColor/:START": fill!.ForegroundColor = ToColor(reader.GetAttribute("rgb")!); break;
                case "styleSheet/fills/fill/patternFill/bgColor/:START": fill!.BackgroundColor = ToColor(reader.GetAttribute("rgb")!); break;

                case "styleSheet/borders/border/:START":
                    border = new();
                    break;

                case "styleSheet/borders/border/:END":
                    borders.Add(border!);
                    break;

                case "styleSheet/borders/border/start/:START":
                case "styleSheet/borders/border/end/:START":
                case "styleSheet/borders/border/top/:START":
                case "styleSheet/borders/border/bottom/:START":
                case "styleSheet/borders/border/diagonal/:START":
                case "styleSheet/borders/border/vertical/:START":
                case "styleSheet/borders/border/horizontal/:START":
                    if (reader.GetAttribute("style") is { } style) borderpr = new() { Style = ToEnumAlias<BorderStyles>(style)!.Value };
                    break;

                case "styleSheet/borders/border/start/color/:START":
                case "styleSheet/borders/border/end/color/:START":
                case "styleSheet/borders/border/top/color/:START":
                case "styleSheet/borders/border/bottom/color/:START":
                case "styleSheet/borders/border/diagonal/color/:START":
                case "styleSheet/borders/border/vertical/color/:START":
                case "styleSheet/borders/border/horizontal/color/:START":
                    borderpr!.Color = ToColor(reader.GetAttribute("rgb")!);
                    break;

                case "styleSheet/borders/border/start/:END": border!.Start = borderpr; break;
                case "styleSheet/borders/border/end/:END": border!.End = borderpr; break;
                case "styleSheet/borders/border/top/:END": border!.Top = borderpr; break;
                case "styleSheet/borders/border/bottom/:END": border!.Bottom = borderpr; break;
                case "styleSheet/borders/border/diagonal/:END": border!.Diagonal = borderpr; break;
                case "styleSheet/borders/border/vertical/:END": border!.Vertical = borderpr; break;
                case "styleSheet/borders/border/horizontal/:END": border!.Horizontal = borderpr; break;

                case "styleSheet/cellXfs/xf/:START":
                    cellstyle = new();
                    if (reader.GetAttribute("applyFont") is { } applyFont && ToBool(applyFont)) cellstyle.Font = fonts[ToInt(reader.GetAttribute("fontId")!)];
                    if (reader.GetAttribute("applyFill") is { } applyFill && ToBool(applyFill)) cellstyle.Fill = fills[ToInt(reader.GetAttribute("fillId")!)];
                    if (reader.GetAttribute("applyBorder") is { } applyBorder && ToBool(applyBorder)) cellstyle.Border = borders[ToInt(reader.GetAttribute("borderId")!)];
                    if (reader.GetAttribute("applyNumberFormat") is { } applyNumberFormat && ToBool(applyNumberFormat))
                    {
                        var numFmtId = ToInt(reader.GetAttribute("numFmtId")!);
                        cellstyle.NumberFormat = formats.TryGetValue(numFmtId, out var value)
                            ? value
                            : new NumberFormatId { FormatId = ToEnum<NumberFormats>(numFmtId) };
                    }
                    hasAlignment = reader.GetAttribute("applyAlignment") is { } applyAlignment && ToBool(applyAlignment);
                    break;

                case "styleSheet/cellXfs/xf/:END":
                    cellstyles.Add(cellstyle!);
                    break;

                case "styleSheet/cellXfs/xf/alignment/:START":
                    if (hasAlignment)
                    {
                        Alignment alignment = new();
                        if (reader.GetAttribute("horizontal") is { } horizontal) alignment.HorizontalAlignment = ToEnumAlias<HorizontalAlignmentTypes>(horizontal);
                        if (reader.GetAttribute("indent") is { } indent) alignment.Indent = ToUInt(indent);
                        if (reader.GetAttribute("justifyLastLine") is { } justifyLastLine) alignment.JustifyLastLine = ToBool(justifyLastLine);
                        if (reader.GetAttribute("readingOrder") is { } readingOrder) alignment.ReadingOrder = ToEnum<ReadingOrders>(readingOrder);
                        if (reader.GetAttribute("relativeIndent") is { } relativeIndent) alignment.RelativeIndent = ToInt(relativeIndent);
                        if (reader.GetAttribute("shrinkToFit") is { } shrinkToFit) alignment.ShrinkToFit = ToBool(shrinkToFit);
                        if (reader.GetAttribute("textRotation") is { } textRotation) alignment.TextRotation = ToUInt(textRotation);
                        if (reader.GetAttribute("vertical") is { } vertical) alignment.VerticalAlignment = ToEnumAlias<VerticalAlignmentTypes>(vertical);
                        if (reader.GetAttribute("wrapText") is { } wrapText) alignment.WrapText = ToBool(wrapText);
                        cellstyle!.Alignment = alignment;
                    }
                    break;
            }
        }
        return [.. cellstyles];
    }

    public static void SetStyle(IHaveStyle style, CellStyle cellstyle)
    {
        if (cellstyle.Font is { } font) style.Font = font;
        if (cellstyle.Fill is { } fill) style.Fill = fill;
        if (cellstyle.Border is { } border) style.Border = border;
        if (cellstyle.Alignment is { } alignment) style.Alignment = alignment;
        if (cellstyle.NumberFormat is { } format) style.NumberFormat = format;
    }
}
