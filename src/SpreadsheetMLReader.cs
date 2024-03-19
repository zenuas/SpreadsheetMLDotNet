﻿using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.SharedStringTable;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Worksheets;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace SpreadsheetMLDotNet;

public static class SpreadsheetMLReader
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
        .Where(x => x.Hierarchy.Join("/") == "workbook/sheets/sheet/:START")
        .Select(x => (
                x.Reader.GetAttribute("name")!,
                (
                    x.Reader.GetAttribute("id", FormatNamespaces.StrictRelationship) ??
                    x.Reader.GetAttribute("id", FormatNamespaces.TransitionalRelationship)
                )!
            )
        )
        .ToArray();

    public static Dictionary<int, IStringItem> ReadSharedStrings(Stream shared_strings)
    {
        var sst = new Dictionary<int, IStringItem>();
        RunProperties? rpr = null;
        RichText? rt = null;

        foreach (var (reader, hierarchy) in XmlReader.Create(shared_strings)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "sst/si/t/:TEXT":
                    sst.Add(sst.Count, new Text { Value = reader.Value });
                    break;

                case "sst/si/r/:START":
                    rpr ??= [];
                    rpr.Add(rt = new() { Text = "" });
                    break;

                case "sst/si/r/:END":
                    rt = null;
                    break;

                case "sst/si/r/rPr/name/:START": rt!.FontName = reader.GetAttribute("val")!; break;
                case "sst/si/r/rPr/charset/:START": rt!.CharacterSet = SpreadsheetMLImport.ToEnum<CharacterSets>(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/family/:START": rt!.FontFamily = SpreadsheetMLImport.ToEnum<FontFamilies>(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/b/:START": rt!.Bold = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/i/:START": rt!.Italic = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/strike/:START": rt!.StrikeThrough = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/outline/:START": rt!.Outline = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/shadow/:START": rt!.Shadow = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/condense/:START": rt!.Condense = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/extend/:START": rt!.Extend = SpreadsheetMLImport.ToBool(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/color/:START": rt!.Color = SpreadsheetMLImport.ToColor(reader.GetAttribute("rgb")!); break;
                case "sst/si/r/rPr/sz/:START": rt!.FontSize = SpreadsheetMLImport.ToDouble(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/u/:START": rt!.Underline = SpreadsheetMLImport.ToEnumAlias<UnderlineTypes>(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/vertAlign/:START": rt!.VerticalAlignment = SpreadsheetMLImport.ToEnumAlias<VerticalPositioningLocations>(reader.GetAttribute("val")!); break;
                case "sst/si/r/rPr/scheme/:START": rt!.Scheme = SpreadsheetMLImport.ToEnumAlias<FontSchemeStyles>(reader.GetAttribute("val")!); break;
                case "sst/si/r/t/:TEXT": rt!.Text = reader.Value; break;

                case "sst/si/:END":
                    if (rpr is { } x) sst.Add(sst.Count, x);
                    rpr = null;
                    break;
            }
        }
        return sst;
    }

    public static IEnumerable<(string Cell, object Value)> ReadSheetReader(Stream worksheet, Dictionary<int, IStringItem> shared_strings)
    {
        var cell = "";
        var v = "";
        var t = CellTypes.Number;

        foreach (var (reader, hierarchy) in XmlReader.Create(worksheet)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "worksheet/sheetData/row/c/:START":
                    cell = reader.GetAttribute("r")!;
                    v = "";
                    t = reader.GetAttribute("t") is { } s ? Enums.ParseWithAlias<CellTypes>(s)!.Value : CellTypes.Number;
                    break;

                case "worksheet/sheetData/row/c/v/:TEXT":
                case "worksheet/sheetData/row/c/is/t/:TEXT":
                case "worksheet/sheetData/row/c/is/r/t/:TEXT":
                    v += reader.Value;
                    break;

                case "worksheet/sheetData/row/c/:END":
                    yield return (cell, GetCellValueFormat(v, t, shared_strings));
                    break;
            }
        }
    }

    public static object GetCellValueFormat(string value, CellTypes cell_type, Dictionary<int, IStringItem> shared_strings) => cell_type switch
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
