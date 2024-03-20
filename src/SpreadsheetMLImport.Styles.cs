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
        bool hasAlignment = false;
        var attr = new Dictionary<string, object?>();

        foreach (var (reader, hierarchy) in XmlReader.Create(styles)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "styleSheet/numFmts/numFmt/:START":
                    formats.Add(ToInt(reader.GetAttribute("numFmtId")!), new() { FormatCode = reader.GetAttribute("formatCode")! });
                    break;

                case "styleSheet/fonts/font/:START":
                    attr.Clear();
                    break;

                case "styleSheet/fonts/font/:END":
                    fonts.Add(new()
                    {
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

                case "styleSheet/fonts/font/name/:START": attr["FontName"] = reader.GetAttribute("val")!; break;
                case "styleSheet/fonts/font/charset/:START": attr["CharacterSet"] = ToEnum<CharacterSets>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/family/:START": attr["FontFamily"] = ToEnum<FontFamilies>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/b/:START": attr["Bold"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/i/:START": attr["Italic"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/strike/:START": attr["StrikeThrough"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/outline/:START": attr["Outline"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/shadow/:START": attr["Shadow"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/condense/:START": attr["Condense"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/extend/:START": attr["Extend"] = ToBool(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/color/:START": attr["Color"] = ToColor(reader.GetAttribute("rgb")!); break;
                case "styleSheet/fonts/font/sz/:START": attr["FontSize"] = ToDouble(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/u/:START": attr["Underline"] = ToEnumAlias<UnderlineTypes>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/vertAlign/:START": attr["VerticalAlignment"] = ToEnumAlias<VerticalPositioningLocations>(reader.GetAttribute("val")!); break;
                case "styleSheet/fonts/font/scheme/:START": attr["Scheme"] = ToEnumAlias<FontSchemeStyles>(reader.GetAttribute("val")!); break;

                case "styleSheet/fills/fill/:START":
                    attr.Clear();
                    break;

                case "styleSheet/fills/fill/:END":
                    fills.Add(new()
                    {
                        PatternType = ToEnum<PatternTypes>(attr.GetOrNull("PatternType")),
                        ForegroundColor = ToColor(attr.GetOrNull("ForegroundColor")),
                        BackgroundColor = ToColor(attr.GetOrNull("BackgroundColor")),
                    });
                    break;

                case "styleSheet/fills/fill/patternFill/:START": attr["PatternType"] = ToEnumAlias<PatternTypes>(reader.GetAttribute("patternType")!); break;
                case "styleSheet/fills/fill/patternFill/fgColor/:START": attr["ForegroundColor"] = ToColor(reader.GetAttribute("rgb")!); break;
                case "styleSheet/fills/fill/patternFill/bgColor/:START": attr["BackgroundColor"] = ToColor(reader.GetAttribute("rgb")!); break;

                case "styleSheet/borders/border/:START":
                    attr.Clear();
                    break;

                case "styleSheet/borders/border/:END":
                    BorderPropertiesType? GetBorderPropertiesTypeOrDefault(string name)
                    {
                        return !attr.ContainsKey($"{name}.Style") && !attr.ContainsKey($"{name}.Color")
                            ? null
                            : new()
                            {
                                Style = ToEnum<BorderStyles>(attr.GetOrNull($"{name}.Style")),
                                Color = ToColor(attr.GetOrNull($"{name}.Color")),
                            };
                    }
                    borders.Add(new()
                    {
                        Start = GetBorderPropertiesTypeOrDefault("start"),
                        End = GetBorderPropertiesTypeOrDefault("end"),
                        Top = GetBorderPropertiesTypeOrDefault("top"),
                        Bottom = GetBorderPropertiesTypeOrDefault("bottom"),
                        Diagonal = GetBorderPropertiesTypeOrDefault("diagonal"),
                        Vertical = GetBorderPropertiesTypeOrDefault("vertical"),
                        Horizontal = GetBorderPropertiesTypeOrDefault("horizontal"),
                    });
                    break;

                case "styleSheet/borders/border/start/:START":
                case "styleSheet/borders/border/end/:START":
                case "styleSheet/borders/border/top/:START":
                case "styleSheet/borders/border/bottom/:START":
                case "styleSheet/borders/border/diagonal/:START":
                case "styleSheet/borders/border/vertical/:START":
                case "styleSheet/borders/border/horizontal/:START":
                    attr.Remove($"{hierarchy[3]}.Style");
                    attr.Remove($"{hierarchy[3]}.Color");
                    if (reader.GetAttribute("style") is { } style) attr[$"{hierarchy[3]}.Style"] = ToEnumAlias<BorderStyles>(style);
                    break;

                case "styleSheet/borders/border/start/color/:START":
                case "styleSheet/borders/border/end/color/:START":
                case "styleSheet/borders/border/top/color/:START":
                case "styleSheet/borders/border/bottom/color/:START":
                case "styleSheet/borders/border/diagonal/color/:START":
                case "styleSheet/borders/border/vertical/color/:START":
                case "styleSheet/borders/border/horizontal/color/:START":
                    attr[$"{hierarchy[3]}.Color"] = ToColor(reader.GetAttribute("rgb")!);
                    break;

                case "styleSheet/cellXfs/xf/:START":
                    attr.Clear();
                    if (reader.GetAttribute("applyFont") is { } applyFont && ToBool(applyFont)) attr["Font"] = fonts[ToInt(reader.GetAttribute("fontId")!)];
                    if (reader.GetAttribute("applyFill") is { } applyFill && ToBool(applyFill)) attr["Fill"] = fills[ToInt(reader.GetAttribute("fillId")!)];
                    if (reader.GetAttribute("applyBorder") is { } applyBorder && ToBool(applyBorder)) attr["Border"] = borders[ToInt(reader.GetAttribute("borderId")!)];
                    if (reader.GetAttribute("applyNumberFormat") is { } applyNumberFormat && ToBool(applyNumberFormat))
                    {
                        var numFmtId = ToInt(reader.GetAttribute("numFmtId")!);
                        attr["NumberFormat"] = formats.TryGetValue(numFmtId, out var value)
                            ? value
                            : new NumberFormatId { FormatId = ToEnum<NumberFormats>(numFmtId) };
                    }
                    hasAlignment = reader.GetAttribute("applyAlignment") is { } applyAlignment && ToBool(applyAlignment);
                    break;

                case "styleSheet/cellXfs/xf/:END":
                    cellstyles.Add(new()
                    {
                        Font = ToStruct<Font>(attr.GetOrNull("Font")),
                        Fill = ToStruct<Fill>(attr.GetOrNull("Fill")),
                        Border = ToStruct<Border>(attr.GetOrNull("Border")),
                        NumberFormat = ToObject<INumberFormat>(attr.GetOrNull("NumberFormat")),
                        Alignment = ToStruct<Alignment>(attr.GetOrNull("Alignment")),
                    });
                    break;

                case "styleSheet/cellXfs/xf/alignment/:START":
                    if (hasAlignment)
                    {
                        var alignment = new Dictionary<string, object?>();
                        if (reader.GetAttribute("horizontal") is { } horizontal) alignment["HorizontalAlignment"] = ToEnumAlias<HorizontalAlignmentTypes>(horizontal);
                        if (reader.GetAttribute("indent") is { } indent) alignment["Indent"] = ToUInt(indent);
                        if (reader.GetAttribute("justifyLastLine") is { } justifyLastLine) alignment["JustifyLastLine"] = ToBool(justifyLastLine);
                        if (reader.GetAttribute("readingOrder") is { } readingOrder) alignment["ReadingOrder"] = ToEnum<ReadingOrders>(readingOrder);
                        if (reader.GetAttribute("relativeIndent") is { } relativeIndent) alignment["RelativeIndent"] = ToInt(relativeIndent);
                        if (reader.GetAttribute("shrinkToFit") is { } shrinkToFit) alignment["ShrinkToFit"] = ToBool(shrinkToFit);
                        if (reader.GetAttribute("textRotation") is { } textRotation) alignment["TextRotation"] = ToUInt(textRotation);
                        if (reader.GetAttribute("vertical") is { } vertical) alignment["VerticalAlignment"] = ToEnumAlias<VerticalAlignmentTypes>(vertical);
                        if (reader.GetAttribute("wrapText") is { } wrapText) alignment["WrapText"] = ToBool(wrapText);
                        attr["Alignment"] = new Alignment()
                        {
                            HorizontalAlignment = ToEnum<HorizontalAlignmentTypes>(alignment.GetOrNull("HorizontalAlignment")),
                            Indent = ToUInt(alignment.GetOrNull("Indent")),
                            JustifyLastLine = ToBool(alignment.GetOrNull("JustifyLastLine")),
                            ReadingOrder = ToEnum<ReadingOrders>(alignment.GetOrNull("ReadingOrder")),
                            RelativeIndent = ToInt(alignment.GetOrNull("RelativeIndent")),
                            ShrinkToFit = ToBool(alignment.GetOrNull("ShrinkToFit")),
                            TextRotation = ToUInt(alignment.GetOrNull("TextRotation")),
                            VerticalAlignment = ToEnum<VerticalAlignmentTypes>(alignment.GetOrNull("VerticalAlignment")),
                            WrapText = ToBool(alignment.GetOrNull("WrapText")),
                        };
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
