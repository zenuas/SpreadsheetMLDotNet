namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct NumberFormatCode : INumberFormat
{
    public required string FormatCode { get; init; }
}
