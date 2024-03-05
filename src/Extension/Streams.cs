using Mina.Extension;
using System.IO;

namespace SpreadsheetMLDotNet.Extension;

public static class Streams
{
    public static void WriteLine(this Stream self) => self.Write([(byte)'\r', (byte)'\n']);

    public static void WriteLine(this Stream self, string s)
    {
        self.Write(s);
        self.WriteLine();
    }
}
