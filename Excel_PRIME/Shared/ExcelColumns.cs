using System;
using System.Runtime.CompilerServices;

namespace ExcelPRIME.Shared;

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
    /// Convert ColumnNameRef - Character(s) into a Row - Column Number eg A->1, B->2, AA -> 27
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static (int row, int col, ReadOnlyMemory<char> colName) GetRowColNumbers(this string columnRef)
    {
        if (columnRef.Length == 0)
        {
            return (0, 0, columnRef.AsMemory());
        }

        int col = -1;
        int row = 0;
        int i = 0;
        char c;
        for (; i < columnRef.Length; i++)
        {
            c = columnRef[i];
            int v = c - 'A';
            if ((uint)v < 26u)
            {
                col = ((col + 1) * 26) + v;
            }
            else
            {
                break;
            }
        }

        ReadOnlyMemory<char> colName = columnRef.AsMemory(0, i);
        for (; i < columnRef.Length; i++)
        {
            c = columnRef[i];
            int v = c - '0';
            if ((uint)v >= 10u)
            {
                return (0, col, colName);
            }
            row = (row * 10) + v;
        }
        return (row - 1, col, colName);
    }

}
