﻿using Mina.Attributes;
using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Styles;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
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
    public static void DoExport(Stream stream, Workbook workbook, Dictionary<string, WorksheetCalculation> calc, bool leave_open, FormatNamespace format)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leave_open);
        var reletionship_to_id = new Dictionary<IRelationshipable, string>();

        var cellstyles = WriteWorkbook(zip, workbook, format, calc);

        zip.CreateEntry("[Content_Types].xml")
            .Open()
            .Using(x => WriteContentTypes(x, cellstyles.Length > 1));

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
                .Then(_ => cellstyles.Length > 1, x => x.Concat((new CellStyles(), FormatNamespaces.Styles[(int)format], "styles.xml")))
            ));

        if (cellstyles.Length > 1)
        {
            zip.CreateEntry("xl/styles.xml")
                .Open()
                .Using(x => WriteStyles(x, cellstyles, format));
        }
    }

    public static string AttributesToString(Dictionary<string, string> attr) => attr.Select(kv => $@"{kv.Key}=""{SecurityElement.Escape(kv.Value)}""").Join(" ");
    public static string Indent(int n) => "".PadLeft(n, ' ');

    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, string value) => (value != "").Return(x => { if (x) attr[name] = ToAttribute(value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, int? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, uint? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, double? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, bool? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, Color? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute(Dictionary<string, string> attr, string name, DateTime? value) => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttribute<E>(Dictionary<string, string> attr, string name, E? value) where E : struct, Enum => (value is { }).Return(x => { if (x) attr[name] = ToAttribute(value!.Value); });
    public static bool TryAddAttributeEnumAlias<E>(Dictionary<string, string> attr, string name, E? value) where E : struct, Enum => (value is { }).Return(x => { if (x) attr[name] = ToAttributeEnumAlias(value!.Value); });

    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, string value) => (value != "").Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value)}"/>"""); });
    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, int? value) => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, uint? value) => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, double? value) => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, bool? value) => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, Color? value) => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElement(Stream stream, int indent, string tag, string attr, DateTime? value) => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElement<E>(Stream stream, int indent, string tag, string attr, E? value) where E : struct, Enum => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttribute(value!.Value)}"/>"""); });
    public static bool TryWriteElementEnumAlias<E>(Stream stream, int indent, string tag, string attr, E? value) where E : struct, Enum => (value is { }).Return(x => { if (x) stream.WriteLine($"""{Indent(indent)}<{tag} {attr}="{ToAttributeEnumAlias(value!.Value)}"/>"""); });

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
