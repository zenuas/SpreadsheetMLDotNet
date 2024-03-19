using Mina.Extension;
using SpreadsheetMLDotNet.Data.Styles;
using System.Collections.Generic;
using System.Linq;

namespace SpreadsheetMLDotNet.Data.SharedStringTable;

public class RunProperties : List<RichText>, IStringItem
{
    public override string ToString() => this.Select(x => x.Text).Join("");
}
