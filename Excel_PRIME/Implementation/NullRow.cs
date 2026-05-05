using System;
using System.Collections.Generic;
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

    public ValueTask<IReadOnlyList<ICell?>?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default) => ValueTask.FromResult<IReadOnlyList<ICell?>?>(null);

    public IReadOnlyList<ICell?>? GetAllCells(CancellationToken ct = default) => null;

    public ValueTask<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default) => ValueTask.FromResult<ICell?>(null);

    public ICell? GetCell(int excelColumnIndex, CancellationToken ct = default) => null;

    public ValueTask<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default) => ValueTask.FromResult<ICell?>(null);

    public ICell? GetCell(string columnLetters, CancellationToken ct = default) => null;

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
