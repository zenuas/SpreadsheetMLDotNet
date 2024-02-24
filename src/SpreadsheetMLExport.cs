using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLExport
{
    public static void DoExport(Stream stream, Workbook workbook, bool leave_open, FormatNamespace format)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leave_open);
        var reletionship_to_id = new Dictionary<IRelationshipable, string>();

        zip.CreateEntry("[Content_Types].xml")
            .Open()
            .Using(WriteContentTypes);

        zip.CreateEntry("_rels/.rels")
            .Open()
            .Using(x => WriteRelationships(x, reletionship_to_id, (workbook, FormatNamespaces.OfficeDocuments[(int)format], "xl/workbook.xml")));

        WriteWorkbook(zip, workbook, format, reletionship_to_id);
    }

    public static void WriteContentTypes(Stream stream) => stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
</Types>
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

    public static void WriteWorkbook(ZipArchive zip, Workbook workbook, FormatNamespace format, Dictionary<IRelationshipable, string> reletionship_to_id)
    {
        zip.CreateEntry("xl/workbook.xml")
            .Open()
            .Using(x => WriteWorkbook(x, workbook, format, reletionship_to_id));

        zip.CreateEntry("xl/_rels/workbook.xml.rels")
            .Open()
            .Using(x => WriteRelationships(x, reletionship_to_id, workbook.Worksheets.Select((w, i) => (w.Cast<IRelationshipable>(), FormatNamespaces.Worksheets[(int)format], $"worksheets/sheet{i + 1}.xml"))));

        workbook.Worksheets.Each((worksheet, i) => zip
            .CreateEntry($"xl/worksheets/sheet{i + 1}.xml")
            .Open()
            .Using(x => WriteWorksheet(x, worksheet, format)));
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

    public static void WriteWorksheet(Stream stream, Worksheet worksheet, FormatNamespace format)
    {
        stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""{FormatNamespaces.SpreadsheetMLMains[(int)format]}"">
  <dimension ref=""A1""/>
  <sheetViews>
    <sheetView tabSelected=""1"" workbookViewId=""0""/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight=""13"" />
  <sheetData>
    <row r=""2"" spans=""2:6"">
      <c r=""B2"" t=""str"">
        <v>abcあいう</v>
      </c>
    </row>
  </sheetData>
  <phoneticPr fontId=""1""/>
  <pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>
</worksheet>
");
    }

    public static string AddRelationship(Dictionary<IRelationshipable, string> reletionship_to_id, IRelationshipable reletionship) => reletionship_to_id.GetOrNew(reletionship, () => $"rId{reletionship_to_id.Count + 1}");
}
