using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace ExcelPRIME.Implementation;

// 10.5
// | AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 2.02x slower | 529000.0000 |  58000.0000 |   6000.0000 | 4188.77 MB | 12.51x more |
// | AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 2.05x slower | 520000.0000 |  56000.0000 |   6000.0000 | 4112.05 MB | 12.28x more |
public class ConcurrentXmlNameTable : XmlNameTable
{
    // Thread-safe dictionary to store atomized strings
    private readonly ConcurrentDictionary<string, string> _nameTable;

    public ConcurrentXmlNameTable(IReadOnlyCollection<string> initialNames)
    {
        List<KeyValuePair<string, string>> kvps = new List<KeyValuePair<string, string>>(initialNames.Count);
        kvps.AddRange(initialNames.Select(name => new KeyValuePair<string, string>(name, name)));

        // Initialize the dictionary with case-sensitive string comparison
        _nameTable = new(kvps, StringComparer.Ordinal);
    }


    // Adds a string to the name table and returns the atomized version
    public override string Add(string array)
    {
        //if (array == null) throw new ArgumentNullException(nameof(array));
        // Use GetOrAdd to ensure thread-safe addition and retrieval
        return _nameTable.GetOrAdd(array, key => key);
    }
    // Adds a substring (from a char array) to the name table and returns the atomized version
    public override string Add(char[] array, int offset, int length)
    {
        //if (array == null) throw new ArgumentNullException(nameof(array));
        //if (offset < 0 || length < 0 || offset + length > array.Length)
        //    throw new ArgumentOutOfRangeException();
        // Create a string from the specified range of the char array
        string key = new string(array, offset, length);
        return Add(key);
    }
    // Retrieves the atomized version of a string if it exists in the name table
    public override string? Get(string array)
    {
        //if (array == null) throw new ArgumentNullException(nameof(array));
        // Try to get the atomized string
        _nameTable.TryGetValue(array, out string? existing);
        return existing;
    }
    // Retrieves the atomized version of a substring if it exists in the name table
    public override string? Get(char[] array, int offset, int length)
    {
        //if (array == null) throw new ArgumentNullException(nameof(array));
        //if (offset < 0 || length < 0 || offset + length > array.Length)
        //    throw new ArgumentOutOfRangeException();
        // Create a string from the specified range of the char array
        string key = new string(array, offset, length);
        return Get(key);
    }

    public void Clear()
    {
        _nameTable.Clear();
    }
}

