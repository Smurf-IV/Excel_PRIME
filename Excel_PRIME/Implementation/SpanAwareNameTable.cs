using System;
using System.Collections.Generic;
using System.Xml;

namespace ExcelPRIME.Implementation;

// A custom IEqualityComparer for strings that can take a span for lookup.
public class SpanAwareStringEqualityComparer : EqualityComparer<string>
{
    // The standard string equality comparer.
    private readonly StringComparer _stringComparer;

    public SpanAwareStringEqualityComparer(StringComparer stringComparer)
    {
        _stringComparer = stringComparer ?? throw new ArgumentNullException(nameof(stringComparer));
    }

    public override bool Equals(string? x, string? y) =>
        // For normal string-to-string comparisons, fall back to the standard comparer.
        _stringComparer.Equals(x, y);

    public override int GetHashCode(string obj) =>
        // For normal string hash code, use the standard comparer.
        _stringComparer.GetHashCode(obj);

    // This method is for performing span-based lookups and is used internally by the dictionary.
    public bool Equals(ReadOnlySpan<char> x, string y) => x.SequenceEqual(y.AsSpan());

    public int GetHashCode(ReadOnlySpan<char> span)
    {
        var hash = new HashCode();
        foreach (var c in span)
        {
            hash.Add(c);
        }
        return hash.ToHashCode();
    }
}

public class SpanAwareNameTable : XmlNameTable
{
    private readonly NameTable _internalNameTable = new NameTable();
    private readonly Dictionary<string, string> _atomizedStringCache;

    public SpanAwareNameTable()
    {
        // Use a standard dictionary with string keys, but use the custom comparer.
        _atomizedStringCache = new Dictionary<string, string>(new SpanAwareStringEqualityComparer(StringComparer.Ordinal));
    }

    public override string Add(string array) =>
        // This method can still call the internal NameTable directly.
        _internalNameTable.Add(array);

    public override string Add(char[] array, int offset, int length)
    {
        var span = new ReadOnlySpan<char>(array, offset, length);

        // Use the Span-based lookup to find an existing atomized string.
        // We use a custom helper method for the lookup to leverage our comparer.
        if (TryGetValue(span, out var atomizedString))
        {
            return atomizedString;
        }
        else
        {
            // If not found, create a new atomized string and store it.
            atomizedString = _internalNameTable.Add(array, offset, length);
            _atomizedStringCache[atomizedString] = atomizedString; // Add to our cache.
            return atomizedString;
        }
    }

    private bool TryGetValue(ReadOnlySpan<char> span, out string value)
    {
        // Cast the comparer to access the span-based Equals and GetHashCode.
        var comparer = (SpanAwareStringEqualityComparer)_atomizedStringCache.Comparer;

        foreach (var kvp in _atomizedStringCache)
        {
            if (comparer.Equals(span, kvp.Key))
            {
                value = kvp.Value;
                return true;
            }
        }
        value = default;
        return false;
    }

    public override string Get(string array) => _internalNameTable.Get(array);

    public override string Get(char[] array, int offset, int length)
    {
        // Perform a lookup based on the span.
        var span = new ReadOnlySpan<char>(array, offset, length);
        TryGetValue(span, out var atomizedString);
        return atomizedString;
    }
}

/*
 Summary of performance considerations

    No intermediate allocations: For Get operations and for checking if a name already exists in the Add operation, this implementation does not create a new string object. It uses the ReadOnlySpan<char> directly for comparison, which is very fast and avoids garbage collection pressure.
    Direct memory access: By implementing a custom hash table, you have granular control over memory layout and can avoid the overhead of standard library collections.
    Hash code for ReadOnlySpan: It uses System.HashCode to calculate hash codes directly from the span's memory, which is highly efficient.
    Simplified implementation: This example uses a simple open-addressing, linked-list-based collision resolution. A production-ready version might use a more advanced strategy, like double hashing, and implement resizing to handle exceeding capacity more gracefully. 
 */
public class HighPerformanceNameTable : XmlNameTable
{
    // A custom struct to hold references for our atomized strings.
    private struct StringAtom
    {
        public int Hash;
        public string AtomizedString;
        public int Next; // For handling hash collisions
    }

    private StringAtom[] _atoms;
    private int[] _buckets;
    private int _count;
    private readonly int _capacity;

    public HighPerformanceNameTable(int capacity = 1024)
    {
        _capacity = capacity;
        _atoms = new StringAtom[capacity];
        _buckets = new int[capacity];
        for (int i = 0; i < capacity; i++)
        {
            _buckets[i] = -1; // -1 indicates an empty bucket
        }
    }

    public override string Add(string array)
    {
        return Add(array.AsSpan());
    }

    public override string Add(char[] array, int offset, int length)
    {
        return Add(new ReadOnlySpan<char>(array, offset, length));
    }

    private string Add(ReadOnlySpan<char> span)
    {
        int hash = CalculateHashCode(span) & (_capacity - 1); // Simple modulo for bucket index

        // Check if the item already exists using span-based comparisons.
        for (int i = _buckets[hash]; i != -1; i = _atoms[i].Next)
        {
            if (span.SequenceEqual(_atoms[i].AtomizedString.AsSpan()))
            {
                return _atoms[i].AtomizedString;
            }
        }

        // If the table is full, throw an exception or resize (more complex).
        if (_count >= _capacity)
        {
            throw new InvalidOperationException("NameTable capacity exceeded.");
        }

        // Add the new item by allocating a new string.
        string atomizedString = span.ToString();
        int newAtomIndex = _count++;
        _atoms[newAtomIndex].Hash = hash;
        _atoms[newAtomIndex].AtomizedString = atomizedString;
        _atoms[newAtomIndex].Next = _buckets[hash];
        _buckets[hash] = newAtomIndex;

        return atomizedString;
    }

    public override string Get(string array)
    {
        return Get(array.AsSpan());
    }

    public override string Get(char[] array, int offset, int length)
    {
        return Get(new ReadOnlySpan<char>(array, offset, length));
    }

    private string Get(ReadOnlySpan<char> span)
    {
        int hash = CalculateHashCode(span) & (_capacity - 1);

        for (int i = _buckets[hash]; i != -1; i = _atoms[i].Next)
        {
            if (span.SequenceEqual(_atoms[i].AtomizedString.AsSpan()))
            {
                return _atoms[i].AtomizedString;
            }
        }
        return null;
    }

    // A simple, fast hash code calculation for ReadOnlySpan<char>
    private int CalculateHashCode(ReadOnlySpan<char> span)
    {
        var hashCode = new HashCode();
        hashCode.AddBytes(System.Runtime.InteropServices.MemoryMarshal.AsBytes(span));
        return hashCode.ToHashCode();
    }
}
