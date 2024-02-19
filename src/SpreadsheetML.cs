using Mina.Extension;
using SpreadsheetMLReader.Data;
using SpreadsheetMLReader.Extension;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace SpreadsheetMLReader;

public static class SpreadsheetML
{
    public static WorkbookReader CreateWorkbookReader(string xlsx_path) => CreateWorkbookReader(File.OpenRead(xlsx_path));

    public static WorkbookReader CreateWorkbookReader(Stream stream)
    {
        var zip = new ZipArchive(stream, ZipArchiveMode.Read, true);
        var entries = zip.GetEntries();
        var workbook_path = entries["_rels/.rels"]
            .Open()
            .Using(ReadRelationshipToWorkbookPath);

        var dir = Path.GetDirectoryName(workbook_path)!;
        var (_, shared_strings_path, id_to_sheetpath) = entries[ZipArchives.ConvertZipPath(Path.Combine(dir, "_rels", Path.GetFileName(workbook_path) + ".rels"))]
            .Open()
            .Using(x => ReadWorkbookRelationshipToFilePath(dir, x));

        var shared_strings = shared_strings_path != ""
            ? entries[shared_strings_path].Open().Using(ReadSharedStrings)
            : [];

        var name_ids = entries[workbook_path].Open().Using(ReadSheetNameToId);
        var name_to_ids = name_ids.ToDictionary();

        return new WorkbookReader()
        {
            WorkSheetNames = name_ids.Select(x => x.Name).ToArray(),
            OpenWorksheet = (sheet_name) =>
            {
                var worksheet = entries[id_to_sheetpath[name_to_ids[sheet_name]]].Open();
                return new WorksheetReader()
                {
                    GetEnumerator_ = () => ReadSheetReader(worksheet, shared_strings).GetEnumerator(),
                    Dispose = worksheet.Dispose,
                };
            },
            Dispose = zip.Dispose,
        };
    }

    public static string ReadRelationshipToWorkbookPath(Stream rels) => XmlReader.Create(rels)
        .UsingDefer(x => x.GetIterator())
        .Where(x => x.NodeType == XmlNodeType.Element && x.Name == "Relationship")
        .Where(x => x.GetAttribute("Type")!.In(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"))
        .Select(x => x.GetAttribute("Target")!)
        .First();

    public static (string StylesPath, string SharedStringsPath, Dictionary<string, string> IdToSheetPath) ReadWorkbookRelationshipToFilePath(string dir, Stream workbook_rels)
    {
        var styles_path = "";
        var shared_strings_path = "";
        var id_to_sheetpath = new Dictionary<string, string>();

        XmlReader.Create(workbook_rels)
            .UsingDefer(x => x.GetIterator())
            .Where(x => x.NodeType == XmlNodeType.Element && x.Name == "Relationship")
            .Each(x =>
            {
                var target = ZipArchives.ConvertZipPath(Path.Combine(dir, x.GetAttribute("Target")!));
                switch (x.GetAttribute("Type"))
                {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                    case "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet":
                        id_to_sheetpath.TryAdd(x.GetAttribute("Id")!, target);
                        break;

                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                    case "http://purl.oclc.org/ooxml/officeDocument/relationships/styles":
                        styles_path = target;
                        break;

                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
                    case "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings":
                        shared_strings_path = target;
                        break;
                }
            });
        return (styles_path, shared_strings_path, id_to_sheetpath);
    }

    public static (string Name, string Id)[] ReadSheetNameToId(Stream workbook) => XmlReader.Create(workbook)
        .UsingDefer(x => x.GetIteratorWithHierarchy())
        .Where(x => x.Reader.NodeType == XmlNodeType.Element && x.Hierarchy.Join("/") == "workbook/sheets/sheet")
        .Select(x => (
                x.Reader.GetAttribute("name")!,
                (
                    x.Reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ??
                    x.Reader.GetAttribute("id", "http://purl.oclc.org/ooxml/officeDocument/relationships")
                )!
            )
        )
        .ToArray();

    public static Dictionary<int, string> ReadSharedStrings(Stream shared_strings) => XmlReader.Create(shared_strings)
        .UsingDefer(x => x.GetIteratorWithHierarchy())
        .Where(x => x.Reader.NodeType == XmlNodeType.Text && x.Hierarchy.Join("/") == "sst/si/t")
        .Select((x, i) => (Index: i, x.Reader.Value))
        .ToDictionary(x => x.Index, x => x.Value);

    public static IEnumerable<(string Cell, object Value)> ReadSheetReader(Stream worksheet, Dictionary<int, string> shared_strings)
    {
        var cell = "";
        var v = "";
        var t = CellType.InlineString;
        foreach (var (reader, hierarchy) in XmlReader.Create(worksheet)
            .UsingDefer(x => x.GetIteratorWithHierarchy())
            .Where(x =>
                (x.Reader.NodeType == XmlNodeType.Text && x.Hierarchy.Join("/") == "worksheet/sheetData/row/c/v") ||
                (x.Reader.NodeType.InStruct(XmlNodeType.Element, XmlNodeType.EndElement) && x.Hierarchy.Join("/") == "worksheet/sheetData/row/c"))
            )
        {

            if (reader.NodeType == XmlNodeType.Element)
            {
                cell = reader.GetAttribute("r")!;
                v = "";
                t = reader.GetAttribute("t") is { } s ? Enums.ParseWithAlias<CellType>(s)!.Value : CellType.String;
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                yield return (cell, GetCellValueFormat(v, t, shared_strings));
            }
            else
            {
                v = reader.Value;
            }
        }
    }

    public static object GetCellValueFormat(string value, CellType cell_type, Dictionary<int, string> shared_strings) => cell_type switch
    {
        CellType.Boolean => value == "1",
        CellType.Date => DateTime.TryParse(value, out var date) ? date : "",
        CellType.Error => value,
        CellType.InlineString => value,
        CellType.String => value,
        CellType.Number => double.TryParse(value, out var num) ? num : value,
        CellType.SharedString => shared_strings[int.Parse(value)],
        _ => throw new ArgumentException(nameof(cell_type)),
    };

    public static (int Row, int Column) ConvertCellAddress(string cell)
    {
        var split = cell.IndexOfAny(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']);
        return (int.Parse(cell[split..]), ConvertColumnNameToIndex(cell[0..split]));
    }

    public static string ConvertCellAddress(int row, int col) => $"{ConvertColumnIndexToName(col)}{row}";

    public static int ConvertColumnNameToIndex(string col) => col.Aggregate(0, (n, c) => (n * 26) + c - 'A' + 1);

    public static string ConvertColumnIndexToName(int col) => (col <= 26 ? "" : ConvertColumnIndexToName((col - 1) / 26)) + (char)('A' + ((col - 1) % 26));
}
