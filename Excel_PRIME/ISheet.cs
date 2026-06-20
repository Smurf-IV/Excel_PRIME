using System.Runtime.CompilerServices;
using System.Threading;

namespace ExcelPRIME;


/// <summary>
/// Access contract for sheets
/// </summary>
public interface ISheet : IDisposable
{
    /// <summary>
    /// This Sheets name
    /// </summary>
    string Name { get; }

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
    /// <param name="cellGetMode">How are the cells populated</param>
    /// <param name="ct"></param>
    IEnumerable<IRow?> GetRowData(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="startExcelColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="endExcelColumn">last col ref (start+width)</param>
    /// <param name="ct"></param>
    IEnumerable<Cell[]?> GetRowData(int startRow, int startExcelColumn, int endExcelColumn, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="startExcelColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="endExcelColumn">last col ref</param>
    /// <param name="ct"></param>
    IEnumerable<Cell[]?> GetRowData(int startRow, ReadOnlySpan<char> startExcelColumn, ReadOnlySpan<char> endExcelColumn, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Using $A$1:$A$1 style, to return data from: a single cell, a single column, a matrix / table 
    /// </summary>
    IEnumerable<Cell[]> GetDefinedRange(DefinedRange range, [EnumeratorCancellation] CancellationToken ct = default);
}

/// <summary>
/// Access contract for sheets
/// </summary>
public interface ISheetAsync : ISheet
{
    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="cellGetMode">How are the cells populated</param>
    /// <param name="ct"></param>
    IAsyncEnumerable<IRowAsync?> GetRowDataAsync(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="startExcelColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="endExcelColumn">last col ref (start+width)</param>
    /// <param name="ct"></param>
    IAsyncEnumerable<Cell[]?> GetRowDataAsync(int startRow, int startExcelColumn, int endExcelColumn, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Returns the row data at the current iterated row
    /// </summary>
    /// <param name="startRow">Skip over the headers / blanks etc</param>
    /// <param name="startExcelColumn">start at a certain matrix / table topleft data cell</param>
    /// <param name="endExcelColumn">last col ref</param>
    /// <param name="ct"></param>
    IAsyncEnumerable<Cell[]?> GetRowDataAsync(int startRow, ReadOnlySpan<char> startExcelColumn, ReadOnlySpan<char> endExcelColumn, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Using $A$1:$A$1 style, to return data from: a single cell, a single column, a matrix / table 
    /// </summary>
    IAsyncEnumerable<Cell[]> GetDefinedRangeAsync(DefinedRange range, [EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Asynchronously retrieves the dimensions of the sheet.
    /// </summary>
    /// <param name="ct"></param>
    /// <returns></returns>
    Task<(int Height, int Width)> GetSheetDimensionsAsync(CancellationToken ct = default);
}