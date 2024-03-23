using Mina.Extension;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;

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

        var (book, name_id_states) = entries[workbook_path].Open().Using(ReadWorkbook);
        name_id_states.Each(nameidstate => entries[id_to_sheetpath[nameidstate.Id]]
            .Open()
            .Using(x =>
            {
                var sheet = ReadWorksheet(x, shared_strings, cellstyles);
                sheet.Name = nameidstate.Name;
                sheet.SheetState = nameidstate.SheetState;
                book.Worksheets.Add(sheet);
            }));
        return book;
    }

    public static int ToInt(string value) => int.Parse(value);

    public static uint ToUInt(string value) => uint.Parse(value);

    public static double ToDouble(string value) => double.Parse(value);

    public static bool ToBool(string value) => value == "1";

    public static Color? ToColor(string value) => string.IsNullOrEmpty(value) ? null : Colors.FromStringArgb(value);

    public static DateTime ToDateTime(string value) => DateTime.Parse(value);

    public static T ToEnum<T>(int value) where T : struct => (T)(object)value;

    public static T ToEnum<T>(string value) where T : struct => Enum.Parse<T>(value);

    public static T? ToEnumAlias<T>(string value) where T : struct, Enum => Enums.ParseWithAlias<T>(value);


    public static string ToString(object? value) => value is string x ? x : "";

    public static int? ToInt(object? value) => value is int x ? x : null;

    public static uint? ToUInt(object? value) => value is uint x ? x : null;

    public static double? ToDouble(object? value) => value is double x ? x : null;

    public static bool? ToBool(object? value) => value is bool x ? x : null;

    public static Color? ToColor(object? value) => value is Color x ? x : null;

    public static DateTime? ToDateTime(object? value) => value is DateTime x ? x : null;

    public static T? ToEnum<T>(object? value) where T : struct => ToStruct<T>(value);

    public static T? ToStruct<T>(object? value) where T : struct => value is T x ? x : null;

    public static T? ToObject<T>(object? value) where T : class => value is T x ? x : null;
}
