using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Extension;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
{
    public static bool TryAddStyleIndex(IHaveStyle have, List<CellStyle> cellstyles, out int index)
    {
        index = -1;
        if (have.Font is null &&
            have.Fill is null &&
            have.Border is null &&
            have.Alignment is null &&
            have.NumberFormat is null) return false;

        var style = new CellStyle { Font = have.Font, Fill = have.Fill, Border = have.Border, Alignment = have.Alignment, NumberFormat = have.NumberFormat };
        index = cellstyles.IndexOf(style);
        if (index < 0)
        {
            index = cellstyles.Count;
            cellstyles.Add(style);
        }
        return true;
    }

    public static void WriteStyles(Stream stream, CellStyle[] cellstyles, FormatNamespace format)
    {
        stream.WriteLine($"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{FormatNamespaces.SpreadsheetMLMains[(int)format]}">
""");
        var format_count = 100;
        var format_list = cellstyles.Select(x => x.NumberFormat).OfType<NumberFormatCode>().ToHashSet().OrderBy(x => x.FormatCode).ToList();
        var formats = format_list.ToDictionary(x => x, x => format_count++);
        if (formats.Count > 0)
        {
            stream.WriteLine($"""  <numFmts count="{formats.Count}">""");
            foreach (var fmt in format_list)
            {
                stream.WriteLine($"""    <numFmt numFmtId="{formats[fmt]}" formatCode="{SecurityElement.Escape(fmt.FormatCode)}"/>""");
            }
            stream.WriteLine("  </numFmts>");
        }

        List<Font> fonts = [new(), .. cellstyles.Select(x => x.Font).OfType<Font>().ToHashSet()]; // index 0 cannot be used in Excel
        stream.WriteLine($"""  <fonts count="{fonts.Count}">""");
        foreach (var font in fonts)
        {
            if (!ExistsFontSetting(font))
            {
                stream.WriteLine("    <font/>");
            }
            else
            {
                stream.WriteLine("    <font>");
                WriteFontAttribute(stream, 6, font);
                stream.WriteLine("    </font>");
            }
        }
        stream.WriteLine("  </fonts>");

        List<Fill> fills = [new(), new(), .. cellstyles.Select(x => x.Fill).OfType<Fill>().ToHashSet()]; // index 0 and 1 cannot be used in Excel
        stream.WriteLine($"""  <fills count="{fills.Count}">""");
        foreach (var fill in fills)
        {
            if (fill.ForegroundColor is null && fill.BackgroundColor is null)
            {
                stream.WriteLine("    <fill/>");
            }
            else
            {
                stream.WriteLine("    <fill>");
                stream.WriteLine($"""      <patternFill patternType="{ToAttributeEnumAlias(fill.PatternType)}">""");
                TryWriteElement(stream, 8, "fgColor", "rgb", fill.ForegroundColor);
                TryWriteElement(stream, 8, "bgColor", "rgb", fill.BackgroundColor);
                stream.WriteLine("      </patternFill>");
                stream.WriteLine("    </fill>");
            }
        }
        stream.WriteLine("  </fills>");

        List<Border> borders = [new(), .. cellstyles.Select(x => x.Border).OfType<Border>().ToHashSet()]; // index 0 cannot be used in Excel
        stream.WriteLine($"""  <borders count="{borders.Count}">""");
        static void WriteBorderStyle(Stream stream, string tag, BorderPropertiesType? borderpr)
        {
            if (borderpr is null) return;
            stream.WriteLine($"""      <{tag} style="{ToAttributeEnumAlias(borderpr.Style)}"{(borderpr.Color is null ? "/" : "")}>""");
            if (borderpr.Color is null) return;
            TryWriteElement(stream, 8, "color", "rgb", borderpr.Color);
            stream.WriteLine($"      </{tag}>");
        }
        foreach (var border in borders)
        {
            if (border.Start is null &&
                border.End is null &&
                border.Top is null &&
                border.Bottom is null &&
                border.Diagonal is null &&
                border.Vertical is null &&
                border.Horizontal is null)
            {
                stream.WriteLine("    <border/>");
            }
            else
            {
                stream.WriteLine("    <border>");
                WriteBorderStyle(stream, "start", border.Start);
                WriteBorderStyle(stream, "end", border.End);
                WriteBorderStyle(stream, "top", border.Top);
                WriteBorderStyle(stream, "bottom", border.Bottom);
                WriteBorderStyle(stream, "diagonal", border.Diagonal);
                WriteBorderStyle(stream, "vertical", border.Vertical);
                WriteBorderStyle(stream, "horizontal", border.Horizontal);
                stream.WriteLine("    </border>");
            }
        }
        stream.WriteLine("  </borders>");

        stream.WriteLine("  <cellXfs>");
        foreach (var cellstyle in cellstyles)
        {
            var attr = new Dictionary<string, string>();
            if (cellstyle.Font is { } font) { attr["fontId"] = fonts.IndexOf(font).ToString(); attr["applyFont"] = "1"; }
            if (cellstyle.Fill is { } fill) { attr["fillId"] = fills.IndexOf(fill).ToString(); attr["applyFill"] = "1"; }
            if (cellstyle.Border is { } border) { attr["borderId"] = borders.IndexOf(border).ToString(); attr["applyBorder"] = "1"; }
            if (cellstyle.Alignment is { }) { attr["applyAlignment"] = "1"; }
            if (cellstyle.NumberFormat is NumberFormatId fmtid) { attr["numFmtId"] = ((int)fmtid.FormatId).ToString(); attr["applyNumberFormat"] = "1"; }
            if (cellstyle.NumberFormat is NumberFormatCode fmtcode) { attr["numFmtId"] = formats[fmtcode].ToString(); attr["applyNumberFormat"] = "1"; }

            var xf_attr_s = AttributesToString(attr);
            if (cellstyle.Alignment is { })
            {
                stream.WriteLine($"    <xf {xf_attr_s}>");
                if (cellstyle.Alignment is { } align)
                {
                    var alignment_attr = new Dictionary<string, string>();
                    TryAddAttributeEnumAlias(alignment_attr, "horizontal", align.HorizontalAlignment);
                    TryAddAttribute(alignment_attr, "indent", align.Indent);
                    TryAddAttribute(alignment_attr, "justifyLastLine", align.JustifyLastLine);
                    TryAddAttribute(alignment_attr, "readingOrder", align.ReadingOrder);
                    TryAddAttribute(alignment_attr, "relativeIndent", align.RelativeIndent);
                    TryAddAttribute(alignment_attr, "shrinkToFit", align.ShrinkToFit);
                    TryAddAttribute(alignment_attr, "textRotation", align.TextRotation);
                    TryAddAttributeEnumAlias(alignment_attr, "vertical", align.VerticalAlignment);
                    TryAddAttribute(alignment_attr, "wrapText", align.WrapText);
                    stream.WriteLine($"      <alignment {AttributesToString(alignment_attr)}/>");
                }
                stream.WriteLine("    </xf>");
            }
            else
            {
                stream.WriteLine($"    <xf {xf_attr_s}/>");
            }
        }
        stream.WriteLine("  </cellXfs>");

        stream.WriteLine("</styleSheet>");
    }

    public static bool ExistsFontSetting(IFont font) =>
        font.FontName != "" ||
        font.CharacterSet is not null ||
        font.FontFamily is not null ||
        font.Bold is not null ||
        font.Italic is not null ||
        font.StrikeThrough is not null ||
        font.Outline is not null ||
        font.Shadow is not null ||
        font.Condense is not null ||
        font.Extend is not null ||
        font.Color is not null ||
        font.FontSize is not null ||
        font.Underline is not null ||
        font.VerticalAlignment is not null ||
        font.Scheme is not null;

    public static void WriteFontAttribute(Stream stream, int indent, IFont font)
    {
        TryWriteElement(stream, indent, "name", "val", font.FontName);
        TryWriteElement(stream, indent, "charset", "val", font.CharacterSet);
        TryWriteElement(stream, indent, "family", "val", font.FontFamily);
        TryWriteElement(stream, indent, "b", "val", font.Bold);
        TryWriteElement(stream, indent, "i", "val", font.Italic);
        TryWriteElement(stream, indent, "strike", "val", font.StrikeThrough);
        TryWriteElement(stream, indent, "outline", "val", font.Outline);
        TryWriteElement(stream, indent, "shadow", "val", font.Shadow);
        TryWriteElement(stream, indent, "condense", "val", font.Condense);
        TryWriteElement(stream, indent, "extend", "val", font.Extend);
        TryWriteElement(stream, indent, "color", "rgb", font.Color);
        TryWriteElement(stream, indent, "sz", "val", font.FontSize);
        TryWriteElementEnumAlias(stream, indent, "u", "val", font.Underline);
        TryWriteElementEnumAlias(stream, indent, "vertAlign", "val", font.VerticalAlignment);
        TryWriteElementEnumAlias(stream, indent, "scheme", "val", font.Scheme);
    }
}
