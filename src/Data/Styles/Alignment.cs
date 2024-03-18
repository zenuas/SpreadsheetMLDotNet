using System;

namespace SpreadsheetMLDotNet.Data.Styles;

public class Alignment : IEquatable<Alignment>
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

    public override bool Equals(object? obj) => Equals(obj as Alignment);

    public bool Equals(Alignment? other) =>
        other is not null &&
        HorizontalAlignment == other.HorizontalAlignment &&
        Indent == other.Indent &&
        JustifyLastLine == other.JustifyLastLine &&
        ReadingOrder == other.ReadingOrder &&
        RelativeIndent == other.RelativeIndent &&
        ShrinkToFit == other.ShrinkToFit &&
        TextRotation == other.TextRotation &&
        VerticalAlignment == other.VerticalAlignment &&
        WrapText == other.WrapText;

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(HorizontalAlignment);
        hash.Add(Indent);
        hash.Add(JustifyLastLine);
        hash.Add(ReadingOrder);
        hash.Add(RelativeIndent);
        hash.Add(ShrinkToFit);
        hash.Add(TextRotation);
        hash.Add(VerticalAlignment);
        hash.Add(WrapText);
        return hash.ToHashCode();
    }
}
