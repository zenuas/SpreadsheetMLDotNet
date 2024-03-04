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

    public bool Equals(Alignment? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        if (!Equals(HorizontalAlignment, other.HorizontalAlignment)) return false;
        if (!Equals(Indent, other.Indent)) return false;
        if (!Equals(JustifyLastLine, other.JustifyLastLine)) return false;
        if (!Equals(ReadingOrder, other.ReadingOrder)) return false;
        if (!Equals(RelativeIndent, other.RelativeIndent)) return false;
        if (!Equals(ShrinkToFit, other.ShrinkToFit)) return false;
        if (!Equals(TextRotation, other.TextRotation)) return false;
        if (!Equals(VerticalAlignment, other.VerticalAlignment)) return false;
        if (!Equals(WrapText, other.WrapText)) return false;
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as Alignment);

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
