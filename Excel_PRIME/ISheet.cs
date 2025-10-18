using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

public interface ISheet : IDisposable
{
    /// <summary>
    /// This Sheets name
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Excel Index of this worksheet (Starts at 1)
    /// </summary>
    int Index { get; }

    /// <summary>
    /// What are the Max dimension defined [Excel Rows, Excel Cells] (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// The Current row iterator offset (Starts at 1)
    /// </summary>
    int CurrentRow { get; }

    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="ct"></param>
    IAsyncEnumerable<IRow?> GetRowDataAsync(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default);
    IEnumerable<IRow?> GetRowData(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="startColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="numberOfColumns">matrix / table width</param>
    /// <param name="ct"></param>
    IAsyncEnumerable<IRow?> GetRowDataAsync(int startRow, int startColumn, int numberOfColumns, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Using A1:A1 style, to return data from: a single cell, a single column, a matrix / table 
    /// </summary>
    IAsyncEnumerable<ICell?[]> GetDefinedRangeAsync(string range, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Retrieves (If exists) the cell data value
    /// </summary>
    Task<ICell?> GetRangeCellAsync(string rangeCell, CancellationToken ct = default);

}