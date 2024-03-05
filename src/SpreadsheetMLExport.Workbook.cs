using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
{
    public static void WriteContentTypes(Stream stream, CellStyles? styles) => stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
{(styles is null ? "" : "  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\r\n")}</Types>
");

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

    public static CellStyle[] WriteWorkbook(ZipArchive zip, Workbook workbook, FormatNamespace format)
    {
        var cellstyles = new List<CellStyle>() { new() }; // index 0 cannot be used in Excel
        workbook.Worksheets.Each((worksheet, i) => zip
            .CreateEntry($"xl/worksheets/sheet{i + 1}.xml")
            .Open()
            .Using(x => WriteWorksheet(x, worksheet, format, cellstyles)));
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
}
