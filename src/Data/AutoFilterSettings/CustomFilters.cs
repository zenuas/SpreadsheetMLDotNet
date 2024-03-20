using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class CustomFilters : List<CustomFilter>, IFilterColumns
{
    public bool? And { get; init; }
}
