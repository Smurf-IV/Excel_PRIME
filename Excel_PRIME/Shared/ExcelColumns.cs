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
    /// Convert ColumnNameRef - Character(s) into a Row - Column Excel Number eg A->1, B->2, AA -> 27
    /// </summary>
    public static (int rowExcel, int colExcel, char[] colName) GetRowColNumbers(this string columnRef) =>
           columnRef.Length == 0 
               ? (0, 0, [])
               : columnRef.AsSpan().GetRowColNumbers();

       [MethodImpl(MethodImplOptions.AggressiveOptimization)]
       public static (int rowExcel, int colExcel, char[] colName) GetRowColNumbers(this ReadOnlySpan<char> columnRefSpan)
       {
           int colExcel = -1;
           int i = 0;
           for (; i < columnRefSpan.Length; i++)
           {
               ref readonly char c = ref columnRefSpan[i];
               int v = c - 'A';
               if ((uint)v < 26u)
               {
                   colExcel = ((colExcel + 1) * 26) + v;
               }
               else
               {
                   break;
               }
           }

           colExcel++; // Make it into the Excel 1 offset #
           char[] colName = columnRefSpan.Slice(0, i).ToArray();
           int rowExcel = columnRefSpan.Slice(i).IntParse();
           return (rowExcel, colExcel, colName);
       }
}
