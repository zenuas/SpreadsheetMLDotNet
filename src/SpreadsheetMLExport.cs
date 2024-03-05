using Mina.Attributes;
using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Extension;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLExport
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

        zip.CreateEntry("xl/workbook.xml")
            .Open()
            .Using(x => WriteWorkbook(x, workbook, format, reletionship_to_id));

        zip.CreateEntry("xl/_rels/workbook.xml.rels")
            .Open()
            .Using(x => WriteRelationships(x, reletionship_to_id,
                workbook.Worksheets.Select((w, i) => (w.Cast<IRelationshipable>(), FormatNamespaces.Worksheets[(int)format], $"worksheets/sheet{i + 1}.xml"))
                .Then(_ => styles is { }, x => x.Concat((styles!, FormatNamespaces.Styles[(int)format], "styles.xml")))
            ));

        var cellstyles = WriteWorkbook(zip, workbook, format);
        if (styles is { })
        {
            zip.CreateEntry("xl/styles.xml")
                .Open()
                .Using(x => WriteStyles(x, styles, cellstyles, format));
        }
    }

    public static string AttributesToString(Dictionary<string, string> attr) => attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ");

    public static string AddRelationship(Dictionary<IRelationshipable, string> reletionship_to_id, IRelationshipable reletionship) => reletionship_to_id.GetOrNew(reletionship, () => $"rId{reletionship_to_id.Count + 1}");

    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, string value) => (value != "").Return(x => { if (x) attr[name] = ToAttribute(value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, int? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, uint? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, double? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, bool? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, Color? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, DateTime? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute<E>(Dictionary<string, string> attr, string name, E? value) where E : struct, Enum => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttributeEnumAlias<E>(Dictionary<string, string> attr, string name, E? value) where E : struct, Enum => (value is { }).Return(x => { if (x) attr[name] = ToAttributeEnumAlias(value!.Value); });

    public static string ToAttribute(string value) => SecurityElement.Escape(value);
    public static string ToAttribute(int value) => value.ToString();
    public static string ToAttribute(uint value) => value.ToString();
    public static string ToAttribute(double value) => value.ToString();
    public static string ToAttribute(bool value) => value ? "1" : "0";
    public static string ToAttribute(Color value) => value.ToStringArgb();
    public static string ToAttribute(DateTime value) => value.ToString();
    public static string ToAttribute<E>(E value) where E : Enum => ToAttribute(value.Cast<int>());
    public static string ToAttributeEnumAlias<E>(E value) where E : Enum => value.GetAttributeOrDefault<AliasAttribute>()!.Name;
}
