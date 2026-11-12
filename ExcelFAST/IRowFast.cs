using System;
using System.Threading;

using ExcelFAST.Implementation;

using ExcelPRIME;


namespace ExcelFAST;

/// <summary>
/// Declaration of the Methods to allow access to rows and cells
/// </summary>
public interface IRowFast : IDisposable
{
    /// <summary>
    /// The Current row offset (Starts at 1, if valid)
    /// </summary>
    int CurrentRow { get; }

    /// <summary>
    /// Reads the row data at the next iterated row.
    /// Increments currentRow
    /// true for data, false for end of sheet
    /// </summary>
    bool GetNextRowData(CancellationToken ct = default);

    /// <summary>
    /// Returns a cell of this row. The underlying Cell is pooled and 
    /// will be reused on subsequent calls - copy data if you need to retain it.
    /// </summary>
    [return: PooledObject]
    PooledCellView GetCellAt(int excelColumn, CancellationToken ct = default);

    /// <summary>
    /// Returns a cell of this row. The underlying Cell is pooled and 
    /// will be reused on subsequent calls - copy data if you need to retain it.
    /// </summary>
    [return: PooledObject]
    PooledCellView GetCellAt(ReadOnlySpan<char> excelColumn, CancellationToken ct = default);

    /// <summary>
    /// Returns a segment of cells for this row. The underlying array is pooled and 
    /// will be reused on subsequent calls - copy data if you need to retain it.
    /// </summary>
    [return: PooledObject]
    ReadOnlySpan<Cell> GetAllCells(CancellationToken ct = default);

    /// <summary>
    /// Returns a segment of cells for this row. The underlying array is pooled and 
    /// will be reused on subsequent calls - copy data if you need to retain it.
    /// </summary>
    /// <param name="startExcelColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="endExcelColumn">last col ref (start+width)</param>
    /// <param name="ct"></param>
    [return: PooledObject]
    ReadOnlySpan<Cell> GetCells(int startExcelColumn, int endExcelColumn, CancellationToken ct = default);

    /// <summary>
    /// Returns a segment of cells for this row. The underlying array is pooled and 
    /// will be reused on subsequent calls - copy data if you need to retain it.
    /// </summary>
    /// <param name="startExcelColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="endExcelColumn">last col ref</param>
    /// <param name="ct"></param>
    [return: PooledObject]
    ReadOnlySpan<Cell> GetCells(ReadOnlySpan<char> startExcelColumn, ReadOnlySpan<char> endExcelColumn, CancellationToken ct = default);
}
