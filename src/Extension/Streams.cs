using System.IO;
using System.Text;

namespace SpreadsheetMLDotNet.Extension;

public static class Streams
{
    public static void WriteLine(this Stream self) => self.Write(Encoding.UTF8.GetBytes("\r\n"));

    public static void WriteLine(this Stream self, string s)
    {
        self.Write(Encoding.UTF8.GetBytes(s));
        self.WriteLine();
    }
}
