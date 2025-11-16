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

/// <summary>
/// Sheet internal access contract
/// </summary>
public interface IXmlSheetReader : IDisposable
{

    /// <summary>
    /// What are the Max dimension defined [Excel Rows, Excel Cells] (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// 
    /// </summary>
    int CurrentRow { get; }

    /// <summary>
    /// Get the row(s), and populate the cells via `cellGetMode`
    /// </summary>
    Task<IRow?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default);

    /// <summary>
    /// Get the row(s), and populate the cells via `cellGetMode`
    /// </summary>
    IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default);

    /// <summary>
    /// TBD (Read the spec!)
    /// </summary>
    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct);
}
