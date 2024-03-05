using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Extension;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
{
    public static bool TryAddStyleIndex(IHaveStyle have, List<CellStyle> cellstyles, out int index)
    {
        index = -1;
        if (have.Font is null &&
            have.Fill is null &&
            have.Border is null &&
            have.Alignment is null) return false;

        var style = new CellStyle() { Font = have.Font, Fill = have.Fill, Border = have.Border, Alignment = have.Alignment };
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
        List<Font> fonts = [new(), .. cellstyles.Select(x => x.Font).OfType<Font>().ToHashSet()]; // index 0 cannot be used in Excel
        stream.WriteLine($"""  <fonts count="{fonts.Count}">""");
        foreach (var font in fonts)
        {
            if (font.FontName == "" &&
                font.CharacterSet is null &&
                font.FontFamily is null &&
                font.Bold is null &&
                font.Italic is null &&
                font.StrikeThrough is null &&
                font.Outline is null &&
                font.Shadow is null &&
                font.Condense is null &&
                font.Extend is null &&
                font.Color is null &&
                font.FontSize is null &&
                font.Underline is null &&
                font.VerticalAlignment is null &&
                font.Scheme is null)
            {
                stream.WriteLine("    <font/>");
            }
            else
            {
                stream.WriteLine("    <font>");
                if (font.FontName != "") stream.WriteLine($"""      <name val="{ToAttribute(font.FontName)}"/>""");
                if (font.CharacterSet is { } charset) stream.WriteLine($"""      <charset val="{ToAttribute(charset)}"/>""");
                if (font.FontFamily is { } family) stream.WriteLine($"""      <family val="{ToAttribute(family)}"/>""");
                if (font.Bold is { } b) stream.WriteLine($"""      <b val="{ToAttribute(b)}"/>""");
                if (font.Italic is { } i) stream.WriteLine($"""      <i val="{ToAttribute(i)}"/>""");
                if (font.StrikeThrough is { } strike) stream.WriteLine($"""      <strike val="{ToAttribute(strike)}"/>""");
                if (font.Outline is { } outline) stream.WriteLine($"""      <outline val="{ToAttribute(outline)}"/>""");
                if (font.Shadow is { } shadow) stream.WriteLine($"""      <shadow val="{ToAttribute(shadow)}"/>""");
                if (font.Condense is { } condense) stream.WriteLine($"""      <condense val="{ToAttribute(condense)}"/>""");
                if (font.Extend is { } extend) stream.WriteLine($"""      <extend val="{ToAttribute(extend)}"/>""");
                if (font.Color is { } color) stream.WriteLine($"""      <color rgb="{ToAttribute(color)}"/>""");
                if (font.FontSize is { } sz) stream.WriteLine($"""      <sz val="{ToAttribute(sz)}"/>""");
                if (font.Underline is { } u) stream.WriteLine($"""      <u val="{ToAttributeEnumAlias(u)}"/>""");
                if (font.VerticalAlignment is { } vertAlign) stream.WriteLine($"""      <vertAlign val="{ToAttributeEnumAlias(vertAlign)}"/>""");
                if (font.Scheme is { } scheme) stream.WriteLine($"""      <scheme val="{ToAttributeEnumAlias(scheme)}"/>""");
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
                if (fill.ForegroundColor is { } fg) stream.WriteLine($"""        <fgColor rgb="{ToAttribute(fg)}"/>""");
                if (fill.BackgroundColor is { } bg) stream.WriteLine($"""        <bgColor rgb="{ToAttribute(bg)}"/>""");
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
            stream.WriteLine($"""        <color rgb="{ToAttribute(borderpr.Color.Value)}"/>""");
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
}
