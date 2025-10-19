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

    ///// <summary>
    ///// Convert ColumnNameRef - Character(s) into a Row - Column Excel Number eg A->1, B->2, AA -> 27
    ///// </summary>
    //[MethodImpl(MethodImplOptions.AggressiveOptimization)]
    //public static (int rowExcel, int colExcel, ReadOnlyMemory<char> colName) GetRowColNumbers(this string columnRef)
    //{
    //    if (columnRef.Length == 0)
    //    {
    //        return (0, 0, columnRef.AsMemory());
    //    }

    //    int colExcel = -1;
    //    int rowExcel = 0;
    //    int i = 0;
    //    char c;
    //    for (; i < columnRef.Length; i++)
    //    {
    //        c = columnRef[i];
    //        int v = c - 'A';
    //        if ((uint)v < 26u)
    //        {
    //            colExcel = ((colExcel + 1) * 26) + v;
    //        }
    //        else
    //        {
    //            break;
    //        }
    //    }

    //    colExcel++; // Make it into the Excel 1 offset #
    //    ReadOnlyMemory<char> colName = columnRef.AsMemory(0, i);
    //    for (; i < columnRef.Length; i++)
    //    {
    //        c = columnRef[i];
    //        int v = c - '0';
    //        if ((uint)v >= 10u)
    //        {
    //            return (0, colExcel, colName);
    //        }
    //        rowExcel = (rowExcel * 10) + v;
    //    }
    //    return (rowExcel, colExcel, colName);
    //}

        /// <summary>
       /// Convert ColumnNameRef - Character(s) into a Row - Column Excel Number eg A->1, B->2, AA -> 27
       /// </summary>
       [MethodImpl(MethodImplOptions.AggressiveOptimization)]
       public static (int rowExcel, int colExcel, ReadOnlyMemory<char> colName) GetRowColNumbers(this string columnRef)
       {
           if (columnRef.Length == 0)
           {
               return (0, 0, ReadOnlyMemory<char>.Empty);
           }

           int colExcel = -1;
           int i = 0;
           ReadOnlySpan<char> columnRefSpan = columnRef.AsSpan();
           for (; i < columnRef.Length; i++)
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
           ReadOnlyMemory<char> colName = columnRef.AsMemory(0, i);
           int rowExcel = columnRefSpan.Slice(i).IntParse();
           return (rowExcel, colExcel, colName);
       }

}
