using System;
using System.Buffers;
using System.Runtime.CompilerServices;

// ReSharper disable ForCanBeConvertedToForeach

namespace ExcelPRIME.FromExternal;

/// <summary>
/// Stolen from here
/// https://stackoverflow.com/a/2652855
/// Then some small modifications for language usage
/// </summary>
internal static class ExcelColumns
{
    private static readonly SearchValues<char> s_asciiLetters = SearchValues.Create("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
    // Precomputed lookup table for first 64 columns (A-)
    private static readonly char[][] s_columnNameCache = new char[64 + 1][];

    static ExcelColumns()
    {
        for (int i = 1; i < s_columnNameCache.Length; i++)
        {
            s_columnNameCache[i] = ComputeColumnName(i);
        }
    }

    // CHANGED: Use lookup table for common columns, reduces allocations
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static char[] GetExcelColumnName(this int columnNumber)
    {
        if (columnNumber > 0 && columnNumber < s_columnNameCache.Length)
        {
            return s_columnNameCache[columnNumber];
        }

        return ComputeColumnName(columnNumber);
    }

    private static char[] ComputeColumnName(int columnNumber)
    {
        // Use stackalloc for temp buffer (max 3 chars for column names up to XFD/16384)
        Span<char> buffer = stackalloc char[4];
        int pos = buffer.Length;
        int dividend = columnNumber;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            buffer[--pos] = (char)(65 + modulo);
            dividend = (dividend - modulo) / 26;
        }

        return buffer.Slice(pos).ToArray();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int ParseColumnOffset(ReadOnlySpan<char> buffer)
    {
        int colExcel = -1;
        int i = buffer.IndexOfAnyExcept(s_asciiLetters);
        if (i == -1) i = buffer.Length;

        for (int j = 0; j < i; j++)
        {
            colExcel = ((colExcel + 1) * 26) + (buffer[j] - 'A');
        }
        return colExcel + 1; // Make it into the Excel 1 offset #
    }


    /// <summary>
    /// Convert ColumnNameRef - Character(s) into a Row - Column Excel Number eg A->1, B->2, AA -> 27
    /// </summary>
    public static (int rowExcel, int colExcel, ReadOnlyMemory<char> colName) GetRowColNumbers(this string columnRef)
        => string.IsNullOrEmpty(columnRef) ? (0, 0, ReadOnlyMemory<char>.Empty) : GetRowColNumbersFromString(columnRef);

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static (int rowExcel, int colExcel, ReadOnlyMemory<char> colName) GetRowColNumbersFromString(string columnRowRef)
    {
        ReadOnlySpan<char> span = columnRowRef.AsSpan();
        int colExcel = -1;
        int i = span.IndexOfAnyExcept(s_asciiLetters);
        if (i == -1) i = span.Length;

        for (int j = 0; j < i; j++)
        {
            colExcel = ((colExcel + 1) * 26) + (span[j] - 'A');
        }

        colExcel++; // Make it into the Excel 1 offset #
        ReadOnlyMemory<char> colName = columnRowRef.AsMemory(0, i);
        int rowExcel = span.Slice(i).IntParse();
        return (rowExcel, colExcel, colName);
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static (int rowExcel, int colExcel, char[] colName) GetRowColNumbers(this ReadOnlySpan<char> columnRowRefSpan)
    {
        int colExcel = -1;
        int i = columnRowRefSpan.IndexOfAnyExcept(s_asciiLetters);
        if (i == -1) i = columnRowRefSpan.Length;

        for (int j = 0; j < i; j++)
        {
            colExcel = ((colExcel + 1) * 26) + (columnRowRefSpan[j] - 'A');
        }

        colExcel++; // Make it into the Excel 1 offset #
        char[] colName = columnRowRefSpan.Slice(0, i).ToArray();
        int rowExcel = columnRowRefSpan.Slice(i).IntParse();
        return (rowExcel, colExcel, colName);
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static int GetColNumber(this ReadOnlySpan<char> columnRefSpan)
    {
        int colExcel = -1;
        int i = columnRefSpan.IndexOfAnyExcept(s_asciiLetters);
        if (i == -1) i = columnRefSpan.Length;

        for (int j = 0; j < i; j++)
        {
            colExcel = ((colExcel + 1) * 26) + (columnRefSpan[j] - 'A');
        }

        return ++colExcel; // Make it into the Excel 1 offset #
    }

}