using SpreadsheetMLDotNet.Data.Address;
using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class AutoFilter
{
    public required IAddressRange Reference { get; init; }
    public List<FilterColumn> FilterColumns { get; init; } = [];
}
