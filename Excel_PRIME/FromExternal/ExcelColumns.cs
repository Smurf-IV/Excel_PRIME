using System;
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
    /// <summary>
    /// Convert Column Number into Column Name - Character(s) eg 1->A, 2->B
    /// </summary>
    /// <param name="columnNumber">Column Number</param>
    /// <returns>Column Name - Character(s)</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static string GetExcelColumnName(this int columnNumber)
    {
        string columnName = string.Empty;   // No need for a StringBuilder,as this will only be done max 3 times

        int dividend = columnNumber;

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = string.Concat(Convert.ToChar(65 + modulo), columnName);
            dividend = (int)((dividend - modulo) / 26);
        }

        return columnName;
    }

    /// <summary>
    /// Convert ColumnNameRef - Character(s) into a Row - Column Excel Number eg A->1, B->2, AA -> 27
    /// </summary>
    public static (int rowExcel, int colExcel, char[] colName) GetRowColNumbers(this string columnRef) =>
           columnRef.Length == 0
               ? (0, 0, [])
               : columnRef.AsSpan().GetRowColNumbers();

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static (int rowExcel, int colExcel, char[] colName) GetRowColNumbers(this ReadOnlySpan<char> columnRowRefSpan)
    {
        int colExcel = -1;
        int i = 0;
        for (; i < columnRowRefSpan.Length; i++)
        {
            ref readonly char c = ref columnRowRefSpan[i];
            if (c >= 'A')
            {
                colExcel = ((colExcel + 1) * 26) + (c - 'A');
            }
            else
            {
                break;
            }
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
        for (int i = 0; i < columnRefSpan.Length; i++)
        {
            ref readonly char c = ref columnRefSpan[i];
            if (c >= 'A')
            {
                colExcel = ((colExcel + 1) * 26) + (c - 'A');
            }
            else
            {
                break;
            }
        }

        return ++colExcel; // Make it into the Excel 1 offset #
    }
}
