using System.Drawing;

namespace SpreadsheetMLDotNet.Extension;

public static class Colors
{
    public static string ToStringArgb(this Color self) => $"{self.A:X2}{self.R:X2}{self.G:X2}{self.B:X2}";
}
