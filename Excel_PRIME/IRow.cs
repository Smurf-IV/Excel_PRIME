using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

/// <summary>
/// Implementation contract for row instances
/// </summary>
public interface IRowBase : IDisposable
{
    /// <summary>
    /// Excel 1 Based
    /// </summary>
    int RowOffset { get; }
}


/// <summary>
/// Implementation contract for row instances
/// </summary>
public interface IRow : IRowBase
{
    /// <summary>
    /// Retrieves _All_ cells from Column 1; through to the width dimension of the sheet
    /// </summary>
    /// <remarks>
    /// Cell 0 will be null, as this is indexing is Excel Based (1 Based)
    /// </remarks>
    ArraySegment<Cell> GetAllCells([EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Retrieves the cell data
    /// </summary>
    Cell GetCell(int excelColumnIndex, CancellationToken ct = default);

    /// <summary>
    /// Retrieves (If exists) the cell data
    /// </summary>
    Cell GetCell(string columnLetters, CancellationToken ct = default);

    /// <summary>
    /// Copies the boxed values of all cells in the row to the specified array.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// Thrown if the <paramref name="values"/> array is <c>null</c>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown if the length of the <paramref name="values"/> array is less than the number of cells in the row.
    /// </exception>
    void CopyBoxedToArray(object?[] values, CancellationToken ct = default);
}

/// <summary>
/// Implementation contract for row instances
/// </summary>
public interface IRowAsync : IRow
{
    /// <summary>
    /// Retrieves _All_ cells within the row, `0` indexed
    /// </summary>
    ValueTask<ArraySegment<Cell>> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default);

    /// <summary>
    /// Retrieves the cell data
    /// </summary>
    ValueTask<Cell> GetCellAsync(int excelColumnIndex, CancellationToken ct = default);

    /// <summary>
    /// Retrieves (If exists) the cell data
    /// </summary>
    ValueTask<Cell> GetCellAsync(string columnLetters, CancellationToken ct = default);
}


/// <summary>
/// Represents a marker interface for a row that contains only null values.
/// </summary>
/// <remarks>
/// Implementations of this interface indicate that all columns in the row are null. This can be used to
/// distinguish between rows with actual data and those that are intentionally empty or uninitialized.
/// </remarks>
public interface INullRow : IRow
{
}

/// <summary>
/// Represents a marker interface for a row that contains only null values.
/// </summary>
/// <remarks>
/// Implementations of this interface indicate that all columns in the row are null. This can be used to
/// distinguish between rows with actual data and those that are intentionally empty or uninitialized.
/// </remarks>
public interface INullRowAsync : INullRow, IRowAsync
{
}