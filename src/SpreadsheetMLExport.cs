using Mina.Attributes;
using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLExport
{
    public static void DoExport(Stream stream, Workbook workbook, bool leave_open, FormatNamespace format)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leave_open);
        var reletionship_to_id = new Dictionary<IRelationshipable, string>();
        var styles = GetStyles(workbook);

        zip.CreateEntry("[Content_Types].xml")
            .Open()
            .Using(x => WriteContentTypes(x, styles));

        zip.CreateEntry("_rels/.rels")
            .Open()
            .Using(x => WriteRelationships(x, reletionship_to_id, (workbook, FormatNamespaces.OfficeDocuments[(int)format], "xl/workbook.xml")));

        var cellstyles = WriteWorkbook(zip, workbook, format, reletionship_to_id, styles);
        if (styles is { })
        {
            zip.CreateEntry("xl/styles.xml")
                .Open()
                .Using(x => WriteStyles(x, styles, cellstyles, format));
        }
    }

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

    public static void WriteContentTypes(Stream stream, CellStyles? styles) => stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
{(styles is null ? "" : "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\r\n")}</Types>
");

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
            stream.Write("      <patternFill patternType=\"solid\">\r\n");
            if (fill.ForegroundColor is { } fg) stream.Write($"        <fgColor rgb=\"{fg.Name.ToUpper()}\"/>\r\n");
            if (fill.BackgroundColor is { } bg) stream.Write($"        <bgColor rgb=\"{bg.Name.ToUpper()}\"/>\r\n");
            stream.Write("      </patternFill>\r\n");
            stream.Write("    </fill>\r\n");
        }
        stream.Write("  </fills>\r\n");

        stream.Write($"  <borders count=\"{styles.Borders.Count}\">\r\n");
        foreach (var borde in styles.Borders)
        {
            stream.Write("    <border/>\r\n");
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

    public static void WriteRelationships(Stream stream, Dictionary<IRelationshipable, string> reletionship_to_id, IEnumerable<(IRelationshipable Reletionship, string Type, string Path)> reletionships) => WriteRelationships(stream, reletionship_to_id, reletionships.ToArray());

    public static void WriteRelationships(Stream stream, Dictionary<IRelationshipable, string> reletionship_to_id, params (IRelationshipable Reletionship, string Type, string Path)[] reletionships)
    {
        stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
");
        reletionships.Each((x, i) => stream.Write($"  <Relationship Id=\"{AddRelationship(reletionship_to_id, x.Reletionship)}\" Type=\"{x.Type}\" Target=\"{x.Path}\"/>\r\n"));
        stream.Write($"</Relationships>\r\n");
    }

    public static CellStyle[] WriteWorkbook(ZipArchive zip, Workbook workbook, FormatNamespace format, Dictionary<IRelationshipable, string> reletionship_to_id, CellStyles? styles)
    {
        zip.CreateEntry("xl/workbook.xml")
            .Open()
            .Using(x => WriteWorkbook(x, workbook, format, reletionship_to_id));

        zip.CreateEntry("xl/_rels/workbook.xml.rels")
            .Open()
            .Using(x => WriteRelationships(x, reletionship_to_id,
                workbook.Worksheets.Select((w, i) => (w.Cast<IRelationshipable>(), FormatNamespaces.Worksheets[(int)format], $"worksheets/sheet{i + 1}.xml"))
                .Then(_ => styles is { }, x => x.Concat((styles!, FormatNamespaces.Styles[(int)format], "styles.xml")))
            ));

        var cellstyles = new List<CellStyle>() { new() };
        workbook.Worksheets.Each((worksheet, i) => zip
            .CreateEntry($"xl/worksheets/sheet{i + 1}.xml")
            .Open()
            .Using(x => WriteWorksheet(x, worksheet, format, styles, cellstyles)));
        return [.. cellstyles];
    }

    public static void WriteWorkbook(Stream stream, Workbook workbook, FormatNamespace format, Dictionary<IRelationshipable, string> reletionship_to_id)
    {
        stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""{FormatNamespaces.SpreadsheetMLMains[(int)format]}"" xmlns:r=""{FormatNamespaces.Relationships[(int)format]}"">
  <sheets>
");
        workbook.Worksheets.Each((x, i) => stream.Write($"    <sheet name=\"{x.Name}\" sheetId=\"{i + 1}\" r:id=\"{AddRelationship(reletionship_to_id, x)}\"/>\r\n"));
        stream.Write(
$@"  </sheets>
</workbook>
");
    }

    public static void WriteWorksheet(Stream stream, Worksheet worksheet, FormatNamespace format, CellStyles? styles, List<CellStyle> cellstyles)
    {
        stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""{FormatNamespaces.SpreadsheetMLMains[(int)format]}"">
  <sheetData>
");
        foreach (var (y, row) in EnumerableRows(worksheet))
        {
            var row_attr = new Dictionary<string, string>();
            if (row.Height is { } height) { row_attr["ht"] = height.ToString(); row_attr["customHeight"] = "1"; }
            if (row.Cells.Count == 0 && row_attr.Count == 0) continue;

            row_attr["r"] = y.ToString();

            stream.Write($"    <row {row_attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ")}>\r\n");
            foreach (var (x, cell) in EnumerableCells(row))
            {
                var cell_attr = new Dictionary<string, string>();
                if (cell.Font is { } || cell.Fill is { } || cell.Border is { })
                {
                    var style = new CellStyle() { Font = cell.Font, Fill = cell.Fill, Border = cell.Border };
                    var styleindex = cellstyles.IndexOf(style);
                    if (styleindex < 0)
                    {
                        styleindex = cellstyles.Count;
                        cellstyles.Add(style);
                    }
                    cell_attr["s"] = styleindex.ToString();
                }
                if (cell.Value is CellValueNull && cell_attr.Count == 0) continue;

                var (cell_type, escaped_value) = GetCellValueFormat(cell.Value);
                cell_attr["r"] = SpreadsheetML.ConvertCellAddress(y, x);
                cell_attr["t"] = cell_type.GetAttributeOrDefault<AliasAttribute>()!.Name;

                stream.Write(
$@"      <c {cell_attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ")}>
        <v>{escaped_value}</v>
      </c>
");
            }
            stream.Write($"    </row>\r\n");
        }
        stream.Write(
$@"  </sheetData>
</worksheet>
");
    }

    public static IEnumerable<(int Index, Row Row)> EnumerableRows(Worksheet worksheet)
    {
        for (var y = 0; y < worksheet.Rows.Count; y++)
        {
            yield return (y + worksheet.StartRowIndex, worksheet.Rows[y]);
        }
    }

    public static IEnumerable<(int Index, Cell Cell)> EnumerableCells(Row row)
    {
        for (var x = 0; x < row.Cells.Count; x++)
        {
            yield return (x + row.StartCellIndex, row.Cells[x]);
        }
    }

    public static IEnumerable<(int Row, int Column, Cell Cell)> EnumerableCells(Workbook workbook)
    {
        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var (y, row) in EnumerableRows(worksheet))
            {
                foreach (var (x, cell) in EnumerableCells(row))
                {
                    yield return (y, x, cell);
                }
            }
        }
    }

    public static (CellTypes CellType, string EscapedValue) GetCellValueFormat(ICellValue value) => value switch
    {
        CellValueBoolean x => (CellTypes.Boolean, x.Value.ToString()),
        CellValueDate x => (CellTypes.Date, x.Value.ToString()),
        CellValueError x => (CellTypes.Error, x.Value.GetAttributeOrDefault<AliasAttribute>()!.Name),
        CellValueDouble x => (CellTypes.Number, x.Value.ToString()),
        CellValueString x => (CellTypes.String, SecurityElement.Escape(x.Value)),
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    public static string AddRelationship(Dictionary<IRelationshipable, string> reletionship_to_id, IRelationshipable reletionship) => reletionship_to_id.GetOrNew(reletionship, () => $"rId{reletionship_to_id.Count + 1}");
}
