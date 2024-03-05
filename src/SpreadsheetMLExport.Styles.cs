using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
{
    public static CellStyles? GetStyles(Workbook workbook)
    {
        var fonts = new HashSet<Font>();
        var fills = new HashSet<Fill>();
        var borders = new HashSet<Border>();
        var alignments = new HashSet<Alignment>();
        foreach (var (_, _, cell) in EnumerableCells(workbook))
        {
            if (cell.Font is { } font) fonts.Add(font);
            if (cell.Fill is { } fill) fills.Add(fill);
            if (cell.Border is { } border) borders.Add(border);
            if (cell.Alignment is { } alignment) alignments.Add(alignment);
        }
        return fonts.Count == 0 && fills.Count == 0 && borders.Count == 0 ? null
            : new()
            {
                Fonts = [new(), .. fonts],        // index 0 cannot be used in Excel
                Fills = [new(), new(), .. fills], // index 0 and 1 cannot be used in Excel
                Borders = [new(), .. borders],    // index 0 cannot be used in Excel
                Alignments = [.. alignments],
            };
    }

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

    public static void WriteStyles(Stream stream, CellStyles styles, CellStyle[] cellstyles, FormatNamespace format)
    {
        stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""{FormatNamespaces.SpreadsheetMLMains[(int)format]}"">
");
        stream.Write($"  <fonts count=\"{styles.Fonts.Count}\">\r\n");
        foreach (var font in styles.Fonts)
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
                stream.Write("    <font/>\r\n");
            }
            else
            {
                stream.Write("    <font>\r\n");
                if (font.FontName != "") stream.Write($"      <name val=\"{ToAttribute(font.FontName)}\"/>\r\n");
                if (font.CharacterSet is { } charset) stream.Write($"      <charset val=\"{ToAttribute(charset)}\"/>\r\n");
                if (font.FontFamily is { } family) stream.Write($"      <family val=\"{ToAttribute(family)}\"/>\r\n");
                if (font.Bold is { } b) stream.Write($"      <b val=\"{ToAttribute(b)}\"/>\r\n");
                if (font.Italic is { } i) stream.Write($"      <i val=\"{ToAttribute(i)}\"/>\r\n");
                if (font.StrikeThrough is { } strike) stream.Write($"      <strike val=\"{ToAttribute(strike)}\"/>\r\n");
                if (font.Outline is { } outline) stream.Write($"      <outline val=\"{ToAttribute(outline)}\"/>\r\n");
                if (font.Shadow is { } shadow) stream.Write($"      <shadow val=\"{ToAttribute(shadow)}\"/>\r\n");
                if (font.Condense is { } condense) stream.Write($"      <condense val=\"{ToAttribute(condense)}\"/>\r\n");
                if (font.Extend is { } extend) stream.Write($"      <extend val=\"{ToAttribute(extend)}\"/>\r\n");
                if (font.Color is { } color) stream.Write($"      <color rgb=\"{ToAttribute(color)}\"/>\r\n");
                if (font.FontSize is { } sz) stream.Write($"      <sz val=\"{ToAttribute(sz)}\"/>\r\n");
                if (font.Underline is { } u) stream.Write($"      <u val=\"{ToAttributeEnumAlias(u)}\"/>\r\n");
                if (font.VerticalAlignment is { } vertAlign) stream.Write($"      <vertAlign val=\"{ToAttributeEnumAlias(vertAlign)}\"/>\r\n");
                if (font.Scheme is { } scheme) stream.Write($"      <scheme val=\"{ToAttributeEnumAlias(scheme)}\"/>\r\n");
                stream.Write("    </font>\r\n");
            }
        }
        stream.Write("  </fonts>\r\n");

        stream.Write($"  <fills count=\"{styles.Fills.Count}\">\r\n");
        foreach (var fill in styles.Fills)
        {
            if (fill.ForegroundColor is null && fill.BackgroundColor is null)
            {
                stream.Write("    <fill/>\r\n");
            }
            else
            {
                stream.Write("    <fill>\r\n");
                stream.Write($"      <patternFill patternType=\"{ToAttributeEnumAlias(fill.PatternType)}\">\r\n");
                if (fill.ForegroundColor is { } fg) stream.Write($"        <fgColor rgb=\"{ToAttribute(fg)}\"/>\r\n");
                if (fill.BackgroundColor is { } bg) stream.Write($"        <bgColor rgb=\"{ToAttribute(bg)}\"/>\r\n");
                stream.Write("      </patternFill>\r\n");
                stream.Write("    </fill>\r\n");
            }
        }
        stream.Write("  </fills>\r\n");

        stream.Write($"  <borders count=\"{styles.Borders.Count}\">\r\n");
        static void WriteBorderStyle(Stream stream, string tag, BorderPropertiesType? borderpr)
        {
            if (borderpr is null) return;
            stream.Write($"      <{tag} style=\"{ToAttributeEnumAlias(borderpr.Style)}\"{(borderpr.Color is null ? "/" : "")}>\r\n");
            if (borderpr.Color is null) return;
            stream.Write($"        <color rgb=\"{ToAttribute(borderpr.Color.Value)}\"/>\r\n");
            stream.Write($"      </{tag}>\r\n");
        }
        foreach (var border in styles.Borders)
        {
            if (border.Start is null &&
                border.End is null &&
                border.Top is null &&
                border.Bottom is null &&
                border.Diagonal is null &&
                border.Vertical is null &&
                border.Horizontal is null)
            {
                stream.Write("    <border/>\r\n");
            }
            else
            {
                stream.Write("    <border>\r\n");
                WriteBorderStyle(stream, "start", border.Start);
                WriteBorderStyle(stream, "end", border.End);
                WriteBorderStyle(stream, "top", border.Top);
                WriteBorderStyle(stream, "bottom", border.Bottom);
                WriteBorderStyle(stream, "diagonal", border.Diagonal);
                WriteBorderStyle(stream, "vertical", border.Vertical);
                WriteBorderStyle(stream, "horizontal", border.Horizontal);
                stream.Write("    </border>\r\n");
            }
        }
        stream.Write("  </borders>\r\n");

        stream.Write("  <cellXfs>\r\n");
        foreach (var cellstyle in cellstyles)
        {
            var attr = new Dictionary<string, string>();
            if (cellstyle.Font is { } font) { attr["fontId"] = styles.Fonts.IndexOf(font).ToString(); attr["applyFont"] = "1"; }
            if (cellstyle.Fill is { } fill) { attr["fillId"] = styles.Fills.IndexOf(fill).ToString(); attr["applyFill"] = "1"; }
            if (cellstyle.Border is { } border) { attr["borderId"] = styles.Borders.IndexOf(border).ToString(); attr["applyBorder"] = "1"; }
            if (cellstyle.Alignment is { }) { attr["applyAlignment"] = "1"; }

            var xf_attr_s = AttributesToString(attr);
            if (cellstyle.Alignment is { })
            {
                stream.Write($"    <xf {xf_attr_s}>\r\n");
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
                    stream.Write($"      <alignment {AttributesToString(alignment_attr)}/>\r\n");
                }
                stream.Write("    </xf>\r\n");
            }
            else
            {
                stream.Write($"    <xf {xf_attr_s}/>\r\n");
            }
        }
        stream.Write("  </cellXfs>\r\n");

        stream.Write("</styleSheet>\r\n");
    }
}
