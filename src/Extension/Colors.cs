using System.Drawing;

namespace SpreadsheetMLDotNet.Extension;

public static class Colors
{
    public static string ToStringArgb(this Color self) => $"{self.A:X2}{self.R:X2}{self.G:X2}{self.B:X2}";

    public static Color FromStringArgb(string hex8) => Color.FromArgb(
        int.Parse(hex8[0..2], System.Globalization.NumberStyles.HexNumber),
        int.Parse(hex8[2..4], System.Globalization.NumberStyles.HexNumber),
        int.Parse(hex8[4..6], System.Globalization.NumberStyles.HexNumber),
        int.Parse(hex8[6..8], System.Globalization.NumberStyles.HexNumber)
    );
}
