namespace SpreadsheetMLDotNet.Data.Styles;

public class Alignment
{
    public HorizontalAlignmentTypes? HorizontalAlignment { get; set; }
    public uint? Indent { get; set; }
    public bool? JustifyLastLine { get; set; }
    public ReadingOrders? ReadingOrder { get; set; }
    public int? RelativeIndent { get; set; }
    public bool? ShrinkToFit { get; set; }
    public uint? TextRotation { get; set; }
    public VerticalAlignmentTypes? VerticalAlignment { get; set; }
    public bool? WrapText { get; set; }
}
