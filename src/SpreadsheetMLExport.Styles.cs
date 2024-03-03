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
        foreach (var (_, _, cell) in EnumerableCells(workbook))
        {
            if (cell.Font is { } font) fonts.Add(font);
            if (cell.Fill is { } fill) fills.Add(fill);
            if (cell.Border is { } border) borders.Add(border);
        }
        return fonts.Count == 0 && fills.Count == 0 && borders.Count == 0 ? null
            : new()
            {
                Fonts = [new(), .. fonts],        // index 0 cannot be used in Excel
                Fills = [new(), new(), .. fills], // index 0 and 1 cannot be used in Excel
                Borders = [new(), .. borders],    // index 0 cannot be used in Excel
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
            stream.Write("    <font/>\r\n");
        }
        stream.Write("  </fonts>\r\n");

        stream.Write($"  <fills count=\"{styles.Fills.Count}\">\r\n");
        foreach (var fill in styles.Fills)
        {
            stream.Write("    <fill>\r\n");
            stream.Write($"      <patternFill patternType=\"{fill.PatternType.GetAttributeOrDefault<AliasAttribute>()!.Name}\">\r\n");
            if (fill.ForegroundColor is { } fg) stream.Write($"        <fgColor rgb=\"{fg.ToStringArgb()}\"/>\r\n");
            if (fill.BackgroundColor is { } bg) stream.Write($"        <bgColor rgb=\"{bg.ToStringArgb()}\"/>\r\n");
            stream.Write("      </patternFill>\r\n");
            stream.Write("    </fill>\r\n");
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
        stream.Write("  </borders>\r\n");

        stream.Write($"  <cellXfs>\r\n");
        foreach (var cellstyle in cellstyles)
        {
            var attr = new Dictionary<string, string>();
            if (cellstyle.Font is { } font) { attr["fontId"] = styles.Fonts.IndexOf(font).ToString(); attr["applyFont"] = "1"; }
            if (cellstyle.Fill is { } fill) { attr["fillId"] = styles.Fills.IndexOf(fill).ToString(); attr["applyFill"] = "1"; }
            if (cellstyle.Border is { } border) { attr["borderId"] = styles.Borders.IndexOf(border).ToString(); attr["applyBorder"] = "1"; }
            stream.Write($"    <xf {attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ")}/>\r\n");
        }
        stream.Write("  </cellXfs>\r\n");

        stream.Write("</styleSheet>\r\n");
    }
}
