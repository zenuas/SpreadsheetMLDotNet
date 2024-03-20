namespace SpreadsheetMLDotNet.Data.Styles;

public readonly struct Alignment
{
    public HorizontalAlignmentTypes? HorizontalAlignment { get; init; }
    public uint? Indent { get; init; }
    public bool? JustifyLastLine { get; init; }
    public ReadingOrders? ReadingOrder { get; init; }
    public int? RelativeIndent { get; init; }
    public bool? ShrinkToFit { get; init; }
    public uint? TextRotation { get; init; }
    public VerticalAlignmentTypes? VerticalAlignment { get; init; }
    public bool? WrapText { get; init; }
}
