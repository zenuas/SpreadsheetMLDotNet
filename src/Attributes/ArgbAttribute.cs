using System;
using System.Drawing;

namespace SpreadsheetMLDotNet.Attributes;

[AttributeUsage(AttributeTargets.Field)]
public class ArgbAttribute(int argb) : Attribute
{
    public Color Color { get; init; } = Color.FromArgb(argb);
}
