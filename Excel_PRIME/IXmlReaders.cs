using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

public interface IXmlWorkBookReader : IDisposable
{
    Task<IReadOnlyDictionary<string, int>> GetSheetNamesAsync(CancellationToken ct);

    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct);
}

public enum RowCellGet
{
    None = 0, // Default: Does not pre get the cells
    PreGet, // If being used in a `ToList` scenario ,then ensure All Cells are got for each iteration
    Background, // TODO: will be used to get the next rows cells after yield this return
}

public interface IXmlSheetReader : IDisposable
{

    /// <summary>
    /// What are the Max dimension defined cells (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }

    int CurrentRow { get; }

    Task<IRow?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default);
    IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default);

    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct);

}
