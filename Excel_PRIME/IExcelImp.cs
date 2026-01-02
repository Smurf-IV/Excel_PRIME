using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using TernaryBool = bool?;

namespace ExcelPRIME;

/// <summary>
/// Internal implementation for handling the specific file types (XLSX / XLSB)
/// </summary>
public interface IExcelImp : IDisposable
{
    /// <summary>
    /// What names exist in this file
    /// </summary>
    IEnumerable<string> SheetNames();

    /// <summary>
    /// Switch functionality to a new sheet
    /// </summary>
    /// <remarks>
    /// `overrideOptionsAndUseSheetOnlyOnce` indicates that:
    /// - `null`:  use the Options value (Default)
    /// - `false`: override and use the OS Temp File (Useful if going to open this again, and it's big)
    /// - `true`:  override and use internal zip rented buffer
    /// </remarks>
    ISheet? GetSheet(string sheetName, TernaryBool overrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default);

    /// <summary>
    /// From the `definedName`s in the xlsx, use the name to return the range data
    /// </summary>
    /// <param name="rangeName"></param>
    /// <param name="useThisSheetName">If passed in, then check that the range exists in that first, before switching to the global name</param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IEnumerable<CellValue?[]> GetDefinedRange(string rangeName, string? useThisSheetName = null,
        [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// From the `definedName`s in the xlsx, use the name to return the range data
    /// </summary>
    /// <param name="rangeName"></param>
    /// <param name="useLocalSheetId">If passed in, then check that the range exists in that first, before switching to the global name</param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IEnumerable<CellValue?[]> GetDefinedRange(string rangeName, int useLocalSheetId,
        [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// User defined range (With or Without `$`'s, e.g., `A1:B2`)
    /// </summary>
    /// <param name="range"></param>
    /// <param name="sheetName"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IEnumerable<CellValue?[]> GetUserRange(string range, string sheetName,
        [EnumeratorCancellation] CancellationToken ct = default);
}

/// <summary>
/// Internal implementation Async for handling the specific file types (XLSX / XLSB)
/// </summary>
public interface IExcelImpAsync : IExcelImp
{
    /// <summary>
    /// Switch functionality to a new sheet
    /// </summary>
    /// <remarks>
    /// `overrideOptionsAndUseSheetOnlyOnce` indicates that:
    /// - `null`:  use the Options value (Default)
    /// - `false`: override and use the OS Temp File (Useful if going to open this again, and it's big)
    /// - `true`:  override and use internal zip rented buffer
    /// </remarks>
    Task<ISheetAsync?> GetSheetAsync(string sheetName, TernaryBool overrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default);

    /// <summary>
    /// From the `definedName`s in the xlsx, use the name to return the range data
    /// </summary>
    /// <param name="rangeName"></param>
    /// <param name="useThisSheetName">If passed in, then check that the range exists in that first, before switching to the global name</param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IAsyncEnumerable<CellValue?[]> GetDefinedRangeAsync(string rangeName, string? useThisSheetName = null,
        [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// User defined range (With or Without `$`'s, e.g., `A1:B2`)
    /// </summary>
    /// <param name="range"></param>
    /// <param name="sheetName"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IAsyncEnumerable<CellValue?[]> GetUserRangeAsync(string range, string sheetName,
        [EnumeratorCancellation] CancellationToken ct = default);

}