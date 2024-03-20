namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct NumberFormatId : INumberFormat
{
    public required NumberFormats FormatId { get; init; }
}
