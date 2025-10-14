using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

public interface IRow : IDisposable
{
    /// <summary>
    /// Excel 1 Based ? TBD
    /// </summary>
    int RowOffset { get; }

    /// <summary>
    /// Retrieves _All_ cells from Column 1; through to the width dimension of the sheet
    /// </summary>
    IAsyncEnumerable<ICell?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default);
    IEnumerable<ICell?> GetAllCells([EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Retrieves the cell data
    /// </summary>
    Task<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default);

    /// <summary>
    /// Retrieves (If exists) the cell data
    /// </summary>
    Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default);

}