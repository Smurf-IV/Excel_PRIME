using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.Implementation;

internal sealed class NullRow(int rowOffset) : IRow
{
    public void Dispose()
    {
    }

    public int RowOffset { get; } = rowOffset;

    public async IAsyncEnumerable<ICell?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        await Task.Yield();
        yield break;
    }

    public IEnumerable<ICell?> GetAllCells(CancellationToken ct = default)
    {
        yield break;
    }

    public Task<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        ICell? nullCell = null;
        return Task.FromResult(nullCell);
    }

    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default)
    {
        ICell? nullCell = null;
        return Task.FromResult(nullCell);
    }
}
