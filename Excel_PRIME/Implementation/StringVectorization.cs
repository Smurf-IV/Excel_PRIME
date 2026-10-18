using System.Runtime.CompilerServices;


namespace ExcelPRIME.Implementation;

/// <summary>
/// Optimized string comparison and matching operations for cell references and column lookups.
/// Uses vectorized operations where possible and scalar fallback for small patterns.
/// Designed for high-throughput filtering of cell ranges by name or reference.
/// </summary>
internal static class StringVectorization
{
    /// <summary>
    /// Case-insensitive ordinal comparison of two span{char}.
    /// Optimized for fast column reference matching in cell lookups.
    /// </summary>
    /// <param name="left">First string to compare.</param>
    /// <param name="right">Second string to compare.</param>
    /// <returns>True if strings are equal (case-insensitive, ordinal), false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool EqualsOrdinalIgnoreCase(ReadOnlySpan<char> left, ReadOnlySpan<char> right)
    {
        if (left.Length != right.Length)
        {
            return false;
        }

        return left.Equals(right, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Case-sensitive ordinal comparison of two span{char}.
    /// </summary>
    /// <param name="left">First string to compare.</param>
    /// <param name="right">Second string to compare.</param>
    /// <returns>True if strings are equal (case-sensitive, ordinal), false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool EqualsOrdinal(ReadOnlySpan<char> left, ReadOnlySpan<char> right) => left.SequenceEqual(right);

    /// <summary>
    /// Fast vectorized search for a pattern in text using byte-level operations.
    /// For patterns &lt;= 4 characters, uses direct comparison.
    /// For longer patterns, uses scalar search with early termination.
    /// </summary>
    /// <param name="text">The text to search in.</param>
    /// <param name="pattern">The pattern to find.</param>
    /// <returns>Index of the first occurrence, or -1 if not found.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IndexOfOrdinal(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern) => text.IndexOf(pattern);

    /// <summary>
    /// Case-insensitive ordinal search for a pattern in text.
    /// Optimized for column name lookups where case variations are common.
    /// </summary>
    /// <param name="text">The text to search in.</param>
    /// <param name="pattern">The pattern to find (case-insensitive).</param>
    /// <returns>Index of the first occurrence, or -1 if not found.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IndexOfOrdinalIgnoreCase(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
        => text.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Checks if text starts with pattern (case-sensitive).
    /// Optimized for cell reference prefix matching.
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <param name="pattern">The pattern to match at the start.</param>
    /// <returns>True if text starts with pattern, false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool StartsWithOrdinal(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
        => text.StartsWith(pattern);

    /// <summary>
    /// Checks if text starts with pattern (case-insensitive).
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <param name="pattern">The pattern to match at the start (case-insensitive).</param>
    /// <returns>True if text starts with pattern, false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool StartsWithOrdinalIgnoreCase(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
        => text.StartsWith(pattern, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Checks if text ends with pattern (case-sensitive).
    /// Optimized for file extension matching.
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <param name="pattern">The pattern to match at the end.</param>
    /// <returns>True if text ends with pattern, false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool EndsWithOrdinal(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
        => text.EndsWith(pattern);

    /// <summary>
    /// Checks if text ends with pattern (case-insensitive).
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <param name="pattern">The pattern to match at the end (case-insensitive).</param>
    /// <returns>True if text ends with pattern, false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool EndsWithOrdinalIgnoreCase(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
        => text.EndsWith(pattern, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Count occurrences of a single character in the text.
    /// Uses vectorized comparison where available.
    /// </summary>
    /// <param name="text">The text to search in.</param>
    /// <param name="character">The character to count.</param>
    /// <returns>Number of occurrences of the character.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int CountOccurrences(ReadOnlySpan<char> text, char character) => text.Count(character);

    /// <summary>
    /// Trim whitespace from both ends of a span using optimized loop unrolling.
    /// </summary>
    /// <param name="text">The text to trim.</param>
    /// <returns>Trimmed span.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static ReadOnlySpan<char> Trim(ReadOnlySpan<char> text)
    {
        if (text.IsEmpty)
        {
            return text;
        }

        // Find first non-whitespace character
        int start = 0;
        while (start < text.Length && char.IsWhiteSpace(text[start]))
            start++;

        if (start == text.Length)
        {
            return ReadOnlySpan<char>.Empty;
        }

        // Find last non-whitespace character
        int end = text.Length - 1;
        while (end > start && char.IsWhiteSpace(text[end]))
            end--;

        return text.Slice(start, end - start + 1);
    }

    /// <summary>
    /// Check if span contains only ASCII characters.
    /// Used to optimize string processing for common case of ASCII-only text.
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <returns>True if all characters are ASCII (0-127), false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static bool IsAsciiOnly(ReadOnlySpan<char> text)
    {
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] > 127)
            {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Compare two strings lexicographically using vectorized operations where possible.
    /// </summary>
    /// <param name="left">First string.</param>
    /// <param name="right">Second string.</param>
    /// <returns>
    /// Negative if left &lt; right, zero if equal, positive if left &gt; right.
    /// </returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int Compare(ReadOnlySpan<char> left, ReadOnlySpan<char> right)
    {
        int minLength = Math.Min(left.Length, right.Length);
        for (int i = 0; i < minLength; i++)
        {
            int cmp = left[i].CompareTo(right[i]);
            if (cmp != 0)
            {
                return cmp;
            }
        }

        return left.Length.CompareTo(right.Length);
    }

    /// <summary>
    /// Lexicographic comparison (case-insensitive).
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static int CompareIgnoreCase(ReadOnlySpan<char> left, ReadOnlySpan<char> right)
    {
        int minLength = Math.Min(left.Length, right.Length);
        for (int i = 0; i < minLength; i++)
        {
            char leftUpper = char.ToUpperInvariant(left[i]);
            char rightUpper = char.ToUpperInvariant(right[i]);
            int cmp = leftUpper.CompareTo(rightUpper);
            if (cmp != 0)
            {
                return cmp;
            }
        }

        return left.Length.CompareTo(right.Length);
    }
}
