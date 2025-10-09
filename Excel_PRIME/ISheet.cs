using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_PRIME;

public interface ISheet : IDisposable
{
    /// <summary>
    /// This Sheets name
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Returns the cell data at the current iterated row
    /// </summary>
    /// <param name="skipRows">Skip over the headers / blanks etc</param>
    /// <param name="startColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="numberOfColumns">matrix / table width</param>
    /// <param name="ct"></param>
    IAsyncEnumerable<object?[]> GetRowDataAsync(int skipRows, int startColumn, int numberOfColumns, CancellationToken ct = default);

    /// <summary>
    /// Retrieves (If exists) the cell data value
    /// </summary>
    Task<object?> GetCellAsync( int rowIndex, int columnIndex, CancellationToken ct = default);

    /// <summary>
    /// Using A1:A1 style, to return data from: a single cell, a single column, a matrix / table 
    /// </summary>
    IAsyncEnumerable<object?[]> GetDefinedRangeAsync(in string range, CancellationToken ct = default);

    /// <summary>
    /// Retrieves (If exists) the cell data value
    /// </summary>
    Task<object?> GetRangeCellAsync(in string rnageCell, CancellationToken ct = default);

}
