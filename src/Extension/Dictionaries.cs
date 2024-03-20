using System.Collections.Generic;

namespace SpreadsheetMLDotNet.Extension;

public static class Dictionaries
{
    public static TValue? GetOrNull<TKey, TValue>(this Dictionary<TKey, TValue> self, TKey key)
        where TKey : notnull
    {
        return self.TryGetValue(key, out var value) ? value : default;
    }
}
