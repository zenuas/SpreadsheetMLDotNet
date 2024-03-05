using System.IO;
using System.Text;

namespace SpreadsheetMLDotNet.Extension;

public static class Streams
{
    public static void WriteLine(this Stream self, string s) => self.Write(Encoding.UTF8.GetBytes($"{s}\r\n"));
}
