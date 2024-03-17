using Mina.Extension;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace SpreadsheetMLDotNet.Extension;

public static class XmlReaders
{
    public static IEnumerable<XmlReader> GetIterator(this XmlReader self) => Lists.Repeat(self).TakeWhile(x => x.Read());

    public static IEnumerable<(XmlReader Reader, IReadOnlyList<string> Hierarchy)> GetIteratorWithHierarchy(this XmlReader self)
    {
        var hierarchy = new List<string>();
        while (self.Read())
        {
            switch (self.NodeType)
            {
                case XmlNodeType.Element:
                    hierarchy.Add(self.Name);
                    yield return (self, [.. hierarchy, ":START"]);
                    if (self.IsEmptyElement) yield return (self, [.. hierarchy, ":END"]);
                    break;

                case XmlNodeType.Text:
                    yield return (self, [.. hierarchy, ":TEXT"]);
                    break;

                case XmlNodeType.EndElement:
                    yield return (self, [.. hierarchy, ":END"]);
                    break;

                default:
                    //yield return (self, hierarchy);
                    break;
            }
            if (self.NodeType == XmlNodeType.EndElement || self.IsEmptyElement) hierarchy.RemoveAt(hierarchy.Count - 1);
        }
    }
}
