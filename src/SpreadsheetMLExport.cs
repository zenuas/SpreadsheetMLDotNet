using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using System.IO;
using System.IO.Compression;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLExport
{
    public static void DoExport(Stream stream, Workbook workbook, bool leave_open, FormatNamespace format)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leave_open);

        zip.CreateEntry("[Content_Types].xml").Open().Using(WriteContentTypes);
        zip.CreateEntry("_rels/.rels").Open().Using(x => WriteworkbookRelationships(x, format));
    }

    public static void WriteContentTypes(Stream stream) => stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
</Types>
");

    public static void WriteworkbookRelationships(Stream stream, FormatNamespace format) => stream.Write(
$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""{FormatNamespaces.Relationships[(int)format]}"">
  <Relationship Id=""rId1"" Type=""{FormatNamespaces.OfficeDocuments[(int)format]}"" Target=""xl/workbook.xml""/>
</Relationships>
");
}
