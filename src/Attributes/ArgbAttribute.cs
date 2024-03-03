using System;
using System.Drawing;

namespace SpreadsheetMLDotNet.Attributes;

[AttributeUsage(AttributeTargets.Field)]
public class ArgbAttribute(uint argb) : Attribute
{
    public Color Color { get; init; } = Color.FromArgb((int)(argb >> 24), (int)((argb >> 16) & 0xFF), (int)((argb >> 8) & 0xFF), (int)(argb & 0xFF));
}
