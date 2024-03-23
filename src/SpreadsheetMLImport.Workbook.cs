using Mina.Extension;
using SpreadsheetMLDotNet.Data;
using SpreadsheetMLDotNet.Data.Workbook;
using SpreadsheetMLDotNet.Data.Worksheets;
using SpreadsheetMLDotNet.Extension;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace SpreadsheetMLDotNet;

public static partial class SpreadsheetMLImport
{
    public static (Workbook Workbook, (string Name, string Id, SheetStates? SheetState)[] NameIdState) ReadWorkbook(Stream workbook)
    {
        var book = new Workbook { Worksheets = [] };
        var name_id_states = new List<(string Name, string Id, SheetStates? SheetState)>();

        foreach (var (reader, hierarchy) in XmlReader.Create(workbook)
            .UsingDefer(x => x.GetIteratorWithHierarchy()))
        {
            switch (hierarchy.Join("/"))
            {
                case "workbook/sheets/sheet/:START":
                    name_id_states.Add((reader.GetAttribute("name")!,
                        (
                            reader.GetAttribute("id", FormatNamespaces.StrictRelationship) ??
                            reader.GetAttribute("id", FormatNamespaces.TransitionalRelationship)
                        )!,
                        ToEnumAlias<SheetStates>(reader.GetAttribute("state") ?? "")));
                    break;
            }
        }
        return (book, [.. name_id_states]);
    }
}
