using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.Implementation;

internal sealed class NullRow(int rowOffset) : IRowAsync
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


    public ICell? GetCell(int excelColumnIndex, CancellationToken ct = default) => null;

    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default)
    {
        ICell? nullCell = null;
        return Task.FromResult(nullCell);
    }

    public ICell? GetCell(string columnLetters, CancellationToken ct = default) => null;

}
