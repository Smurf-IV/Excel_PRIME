using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;


namespace ExcelPRIME.Implementation;

internal sealed class NullRow(int rowOffset) : INullRowAsync
{
    public void Dispose()
    {
    }

    public int RowOffset { get; } = rowOffset;

    public ValueTask<ArraySegment<Cell>> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default) => ValueTask.FromResult<ArraySegment<Cell>>(default);

    public ArraySegment<Cell> GetAllCells(CancellationToken ct = default) => default;

    public ValueTask<Cell> GetCellAsync(int excelColumnIndex, CancellationToken ct = default) => ValueTask.FromResult<Cell>(default);

    public Cell GetCell(int excelColumnIndex, CancellationToken ct = default) => default;

    public ValueTask<Cell> GetCellAsync(string columnLetters, CancellationToken ct = default) => ValueTask.FromResult<Cell>(default);

    public Cell GetCell(string columnLetters, CancellationToken ct = default) => default;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void CopyBoxedToArray(object?[] values, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(values);
        for (int ordinal = 0; ordinal < values.Length; ++ordinal)
        {
            values[ordinal] = null;
        }
    }
}
