using Mina.Extension;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLImport
{
    public static Workbook DoImport(Stream stream, bool leave_open)
    {
        var zip = new ZipArchive(stream, ZipArchiveMode.Read, leave_open);
        var entries = zip.GetEntries();
        var workbook_path = entries["_rels/.rels"]
            .Open()
            .Using(SpreadsheetMLReader.ReadRelationshipToWorkbookPath);

        var dir = Path.GetDirectoryName(workbook_path)!;
        var (styles_path, shared_strings_path, id_to_sheetpath) = entries[ZipArchives.ConvertZipPath(Path.Combine(dir, "_rels", Path.GetFileName(workbook_path) + ".rels"))]
            .Open()
            .Using(x => SpreadsheetMLReader.ReadWorkbookRelationshipToFilePath(dir, x));

        var cellstyles = styles_path != ""
            ? entries[styles_path].Open().Using(ReadStyles)
            : [];

        var shared_strings = shared_strings_path != ""
            ? entries[shared_strings_path].Open().Using(SpreadsheetMLReader.ReadSharedStrings)
            : [];

        var name_ids = entries[workbook_path].Open().Using(SpreadsheetMLReader.ReadSheetNameToId);

        return new Workbook
        {
            Worksheets = name_ids.Select(nameid => entries[id_to_sheetpath[nameid.Id]]
                .Open()
                .Using(x => ReadWorksheet(x, nameid.Name, shared_strings)))
                .ToList()
        };
    }

    public static int ToInt(string value) => int.Parse(value);

    public static uint ToUInt(string value) => uint.Parse(value);

    public static double ToDouble(string value) => double.Parse(value);

    public static bool ToBool(string value) => value == "1";

    public static Color ToColor(string value) => Color.FromName(value);

    public static DateTime ToDateTime(string value) => DateTime.Parse(value);

    public static T ToEnum<T>(int value) where T : struct => (T)(object)value;

    public static T ToEnum<T>(string value) where T : struct => Enum.Parse<T>(value);

    public static T? ToEnumAlias<T>(string value) where T : struct, Enum => Enums.ParseWithAlias<T>(value);

    public static Worksheet ReadWorksheet(Stream worksheet, string sheet_name, Dictionary<int, string> shared_strings)
    {
        var sheet = new Worksheet { Name = sheet_name };
        var cell = "";
        var v = "";
        var t = CellTypes.Number;
        foreach (var (reader, hierarchy) in XmlReader.Create(worksheet)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
        }
        return sheet;
    }
}
