using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace SpreadsheetMLReader.Extension;

public static class ZipArchives
{
    public static Dictionary<string, ZipArchiveEntry> GetEntries(this ZipArchive self) => self.Entries.ToDictionary(x => x.FullName);

    public static string ConvertZipPath(string path) => Path.AltDirectorySeparatorChar != '/' ? path : path.Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
}
