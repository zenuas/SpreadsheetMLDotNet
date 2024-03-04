using Mina.Attributes;
using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Extension;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

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
                if (font.FontName != "") stream.Write($"      <name val=\"{SecurityElement.Escape(font.FontName)}\"/>\r\n");
                if (font.CharacterSet is { } charset) stream.Write($"      <charset val=\"{(int)charset}\"/>\r\n");
                if (font.FontFamily is { } family) stream.Write($"      <family val=\"{(int)family}\"/>\r\n");
                if (font.Bold is { } b) stream.Write($"      <b val=\"{(b ? "1" : "0")}\"/>\r\n");
                if (font.Italic is { } i) stream.Write($"      <i val=\"{(i ? "1" : "0")}\"/>\r\n");
                if (font.StrikeThrough is { } strike) stream.Write($"      <strike val=\"{(strike ? "1" : "0")}\"/>\r\n");
                if (font.Outline is { } outline) stream.Write($"      <outline val=\"{(outline ? "1" : "0")}\"/>\r\n");
                if (font.Shadow is { } shadow) stream.Write($"      <shadow val=\"{(shadow ? "1" : "0")}\"/>\r\n");
                if (font.Condense is { } condense) stream.Write($"      <condense val=\"{(condense ? "1" : "0")}\"/>\r\n");
                if (font.Extend is { } extend) stream.Write($"      <extend val=\"{(extend ? "1" : "0")}\"/>\r\n");
                if (font.Color is { } color) stream.Write($"      <color rgb=\"{color.ToStringArgb()}\"/>\r\n");
                if (font.FontSize is { } sz) stream.Write($"      <sz val=\"{sz}\"/>\r\n");
                if (font.Underline is { } u) stream.Write($"      <u val=\"{u.GetAttributeOrDefault<AliasAttribute>()!.Name}\"/>\r\n");
                if (font.VerticalAlignment is { } vertAlign) stream.Write($"      <vertAlign val=\"{vertAlign.GetAttributeOrDefault<AliasAttribute>()!.Name}\"/>\r\n");
                if (font.Scheme is { } scheme) stream.Write($"      <scheme val=\"{scheme.GetAttributeOrDefault<AliasAttribute>()!.Name}\"/>\r\n");
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
                stream.Write($"      <patternFill patternType=\"{fill.PatternType.GetAttributeOrDefault<AliasAttribute>()!.Name}\">\r\n");
                if (fill.ForegroundColor is { } fg) stream.Write($"        <fgColor rgb=\"{fg.ToStringArgb()}\"/>\r\n");
                if (fill.BackgroundColor is { } bg) stream.Write($"        <bgColor rgb=\"{bg.ToStringArgb()}\"/>\r\n");
                stream.Write("      </patternFill>\r\n");
                stream.Write("    </fill>\r\n");
            }
        }
        stream.Write("  </fills>\r\n");

        stream.Write($"  <borders count=\"{styles.Borders.Count}\">\r\n");
        static void WriteBorderStyle(Stream stream, string tag, BorderPropertiesType? borderpr)
        {
            if (borderpr is null) return;
            stream.Write($"      <{tag} style=\"{borderpr.Style.GetAttributeOrDefault<AliasAttribute>()!.Name}\">\r\n");
            if (borderpr.Color is { } color) stream.Write($"        <color rgb=\"{color.ToStringArgb()}\"/>\r\n");
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

        stream.Write($"  <cellXfs>\r\n");
        foreach (var cellstyle in cellstyles)
        {
            var attr = new Dictionary<string, string>();
            if (cellstyle.Font is { } font) { attr["fontId"] = styles.Fonts.IndexOf(font).ToString(); attr["applyFont"] = "1"; }
            if (cellstyle.Fill is { } fill) { attr["fillId"] = styles.Fills.IndexOf(fill).ToString(); attr["applyFill"] = "1"; }
            if (cellstyle.Border is { } border) { attr["borderId"] = styles.Borders.IndexOf(border).ToString(); attr["applyBorder"] = "1"; }
            if (cellstyle.Alignment is { } alignment) { attr["applyAlignment"] = "1"; }

            if (cellstyle.Alignment is { })
            {
                stream.Write($"    <xf {attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ")}>\r\n");
                if (cellstyle.Alignment is { } align)
                {
                    var alignment_attr = new Dictionary<string, string>();
                    if (align.HorizontalAlignment is { } horizontal) alignment_attr["horizontal"] = horizontal.GetAttributeOrDefault<AliasAttribute>()!.Name;
                    if (align.Indent is { } indent) alignment_attr["indent"] = indent.ToString();
                    if (align.JustifyLastLine is { } justify) alignment_attr["justifyLastLine"] = justify ? "1" : "0";
                    if (align.ReadingOrder is { } reading) alignment_attr["readingOrder"] = ((uint)reading).ToString();
                    if (align.RelativeIndent is { } relative) alignment_attr["relativeIndent"] = relative.ToString();
                    if (align.ShrinkToFit is { } shrink) alignment_attr["shrinkToFit"] = shrink ? "1" : "0";
                    if (align.TextRotation is { } rotation) alignment_attr["textRotation"] = rotation.ToString();
                    if (align.VerticalAlignment is { } vertical) alignment_attr["vertical"] = vertical.GetAttributeOrDefault<AliasAttribute>()!.Name;
                    if (align.WrapText is { } wrap) alignment_attr["wrapText"] = wrap ? "1" : "0";
                    stream.Write($"      <alignment {alignment_attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ")}/>\r\n");
                }
                stream.Write("    </xf>\r\n");
            }
            else
            {
                stream.Write($"    <xf {attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ")}/>\r\n");
            }
        }
        stream.Write("  </cellXfs>\r\n");

        stream.Write("</styleSheet>\r\n");
    }
}
