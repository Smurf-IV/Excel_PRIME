using System;
using System.Text.RegularExpressions;

namespace ExcelPRIME.Shared;

/// <summary>
/// Stolen from here
/// https://stackoverflow.com/a/2652855
/// Then some small modifications for language usage
/// </summary>
internal static partial class ExcelColumns
{
    /// <summary>
    /// Convert Column Number into Column Name - Character(s) eg 1->A, 2->B
    /// </summary>
    /// <param name="columnNumber">Column Number</param>
    /// <returns>Column Name - Character(s)</returns>
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
    /// Covert Column Name - Character(s) into a Column Number eg A->1, B->2, AA -> 27
    /// </summary>
    /// <param name="columnName">Column Name - Character(s)</param>
    /// <param name="includesRowNumber">Has this been stripped, if not - then true</param>
    /// <returns>Column Number</returns>
    public static int GetExcelColumnNumber(this string columnName, bool includesRowNumber = true)
    {
        if (includesRowNumber)
        {
            columnName = RemoveNumbers().Replace(columnName, string.Empty);
        }

        int[] digits = new int[columnName.Length];
        for (int i = 0; i < columnName.Length; ++i)
        {
            digits[i] = Convert.ToInt32(columnName[i]) - 64;
        }
        int mul = 1; int res = 0;
        for (int pos = digits.Length - 1; pos >= 0; --pos)
        {
            res += digits[pos] * mul;
            mul *= 26;
        }
        return res;
    }

    public static int GetRowNumber(this string columnRef)
    {
        if (columnRef.Length == 0)
        {
            return 0;
        }

        int col = -1;
        int row = 0;
        int i = 0;
        char c;
        for (; i < columnRef.Length; i++)
        {
            c = columnRef[i];
            var v = c - 'A';
            if ((uint)v < 26u)
            {
                col = ((col + 1) * 26) + v;
            }
            else
            {
                break;
            }
        }

        for (; i < columnRef.Length; i++)
        {
            c = columnRef[i];
            var v = c - '0';
            if ((uint)v >= 10u)
            {
                return 0;
            }
            row = row * 10 + v;
        }
        return row - 1;
    }

    [GeneratedRegex(@"\d")]
    internal static partial Regex RemoveNumbers();
}
