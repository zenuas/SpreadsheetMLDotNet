using SpreadsheetMLDotNet.Data.Address;

namespace SpreadsheetMLDotNet.Data.AutoFilterSettings;

public class AutoFilter
{
    public required IAddressRange Reference { get; set; }
}
