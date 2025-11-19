using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

/// <summary>
/// Excel file access contract
/// </summary>
public interface IXmlWorkBookReader : IDisposable
{
    /// <summary>
    /// What it says on the tin
    /// </summary>
    IEnumerable<KeyValuePair<string, int>> GetSheetNames(CancellationToken ct);

    /// <summary>
    /// What it says on the tin
    /// </summary>
    IReadOnlyDictionary<string, DefinedRange> GetDefinedRanges(IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct);
}

/// <summary>
/// Excel file access contract
/// </summary>
public interface IXmlWorkBookReaderAsync : IXmlWorkBookReader
{
    /// <summary>
    /// What it says on the tin
    /// </summary>
    IAsyncEnumerable<KeyValuePair<string, int>> GetSheetNamesAsync(CancellationToken ct);

    /// <summary>
    /// What it says on the tin
    /// </summary>
    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(
        IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct);
}

/// <summary>
/// How are the cells retrieved
/// </summary>
public enum RowCellGet
{
    /// <summary>
    /// Default: Does not pre get the cells
    /// </summary>
    None = 0,
    /// <summary>
    /// If being used in a `ToList` scenario ,then ensure All Cells are got for each iteration
    /// </summary>
    PreGet,
    /// <summary>
    /// TODO: will be used to get the next rows cells after yield this return
    /// </summary>
    Background
}
