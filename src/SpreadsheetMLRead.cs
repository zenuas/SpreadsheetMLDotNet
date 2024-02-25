using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLRead
{
    public static WorkbookReader OpenReader(Stream stream, bool leave_open)
    {
        var zip = new ZipArchive(stream, ZipArchiveMode.Read, leave_open);
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
        .Where(x => x.GetAttribute("Type")!.In(FormatNamespaces.OfficeDocuments))
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
                    case FormatNamespaces.StrictWorksheet:
                    case FormatNamespaces.TransitionalWorksheet:
                        id_to_sheetpath.TryAdd(x.GetAttribute("Id")!, target);
                        break;

                    case FormatNamespaces.StrictStyles:
                    case FormatNamespaces.TransitionalStyles:
                        styles_path = target;
                        break;

                    case FormatNamespaces.StrictSharedStrings:
                    case FormatNamespaces.TransitionalSharedStrings:
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
                    x.Reader.GetAttribute("id", FormatNamespaces.StrictRelationship) ??
                    x.Reader.GetAttribute("id", FormatNamespaces.TransitionalRelationship)
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
        var t = CellTypes.Number;
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
                t = reader.GetAttribute("t") is { } s ? Enums.ParseWithAlias<CellTypes>(s)!.Value : CellTypes.Number;
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

    public static object GetCellValueFormat(string value, CellTypes cell_type, Dictionary<int, string> shared_strings) => cell_type switch
    {
        CellTypes.Boolean => value == "1",
        CellTypes.Date => DateTime.TryParse(value, out var date) ? date : "",
        CellTypes.Error => value,
        CellTypes.InlineString => value,
        CellTypes.String => value,
        CellTypes.Number => double.TryParse(value, out var num) ? num : value,
        CellTypes.SharedString => shared_strings[int.Parse(value)],
        _ => throw new ArgumentException(null, nameof(cell_type)),
    };
}
