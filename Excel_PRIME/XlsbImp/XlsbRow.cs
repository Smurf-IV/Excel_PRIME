using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

internal sealed class XlsbRow : IRowAsync
{
    private XlsbStreamReader? _reader;
    private InstanceContext? _instanceContext;
    private int _maxExcelColumnDimension;
    private bool _isDisposed;
    private XlsbCell?[]? _cells;
    private bool _cellsLoaded;

    // Small object pool for Row instances to avoid allocating a new Row per XML row.
    private static readonly ConcurrentBag<XlsbRow> s_pool = new();

    private XlsbRow()
    {
        // Private ctor for pooling. Keep lightweight.
    }

    internal static XlsbRow Rent()
    {
        if (s_pool.TryTake(out XlsbRow? item))
        {
            return item;
        }

        return new XlsbRow();
    }

    private static void Return(XlsbRow row)
    {
        // Reset state so next consumer sees a clean Row.
        row.Reset();
        s_pool.Add(row);
    }

    internal void Initialize(XlsbStreamReader rowElement, InstanceContext instanceContext, int maxColumnDimension)
    {
        _reader = rowElement;
        _instanceContext = instanceContext;
        _maxExcelColumnDimension = maxColumnDimension;
        using (PooledRecordBuffer nextRecord = _reader.ReadNextRecord())
        {
            RowOffset = nextRecord.GetInt32(0) + 1; // Add 1, to resolve back to VBA 1-based index
            //var ifx = nextRecord.GetInt32(4);
            //var flags = nextRecord.GetByte(11);
            //_isRowHidden = (flags & 0x10) != 0;
        }
    }

    private void Reset()
    {
        _isDisposed = false;
        _reader = null;
        _instanceContext = null;
        _maxExcelColumnDimension = 0;
        _cells = null;
        _cellsLoaded = false;
        // Do not reset RowOffset — it will be set again on Initialize
        RowOffset = 0;
    }

    private void DisposeManagedState()
    {
        // Release references to allow GC of contained cells
        _cells = null;
        _cellsLoaded = false;
    }

    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _isDisposed = true;
        DisposeManagedState();

        // Return to pool for reuse
        Return(this);
    }

    /// <InheritDoc />
    public int RowOffset { get; private set; }

    /// <summary>
    /// Ensure cells are read once. Cells are stored in a small array indexed by excel 1-based column offset.
    /// Using an array avoids Dictionary overhead and reduces per-row allocations for typical sheet widths.
    /// </summary>
    internal async Task GetCellsAsync(CancellationToken ct)
    {
        if (_cellsLoaded)
        {
            return;
        }

        if (_reader == null)
        {
            return;
        }
        // Defer allocating the cell array until we actually parse cells to keep Row light-weight when unused.
        XlsbCell?[] localCells = new XlsbCell?[_maxExcelColumnDimension + 1];
        PooledRecordBuffer nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        try
        {
            while (nextRecord.Succeeded
                   && !ct.IsCancellationRequested
                   && nextRecord.RecordType != RecordTypeIdentifier.DATAEND
                   && nextRecord.RecordType != RecordTypeIdentifier.ROWHDR)
            {
                switch (nextRecord.RecordType)
                {
                    case RecordTypeIdentifier.CELLBLANK:
                    case RecordTypeIdentifier.CELLRK:
                    case RecordTypeIdentifier.CELLERROR:
                    case RecordTypeIdentifier.CELLBOOL:
                    case RecordTypeIdentifier.CELLREAL:
                    case RecordTypeIdentifier.CELLST:
                    case RecordTypeIdentifier.CELLISST:
                    case RecordTypeIdentifier.CELLFMLASTRING:
                    case RecordTypeIdentifier.CELLFMLANUM:
                    case RecordTypeIdentifier.CELLFMLABOOL:
                    case RecordTypeIdentifier.CELLFMLAERROR:
                        {
                            XlsbCell? cell = XlsbCell.ConstructCell(nextRecord, _instanceContext!);
                            int offset = cell?.ExcelColumnOffset ?? -1;
                            if (offset > 0 && offset <= _maxExcelColumnDimension)
                            {
                                localCells[offset] = cell;
                            }
                            else
                            {
                                // If the parsed cell offset is outside expected width, skip storing it to avoid OOB.
                                // (Should be rare, defensive programming.)
                            }
                        }
                        break;
                    default:
                        // 37
                        // BrtRwDescent
                        // PRINTWIDTH
                        break;
                }
                nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
            }
        }
        finally
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.ROWHDR)
            {
                _reader.RollBackLastRecord(nextRecord);
            }
            else
            {
                nextRecord.Dispose();
            }
        }
        // publish parsed cells once fully read to avoid partial-visible state
        _cells = localCells;
        _cellsLoaded = true;
    }

    internal void GetCells(CancellationToken ct)
    {
        if (_cellsLoaded)
        {
            return;
        }

        if (_reader == null)
        {
            return;
        }
        // Defer allocating the cell array until we actually parse cells to keep Row light-weight when unused.
        XlsbCell?[] localCells = new XlsbCell?[_maxExcelColumnDimension + 1];
        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        try
        {
            while (nextRecord.Succeeded
                   && !ct.IsCancellationRequested
                   && nextRecord.RecordType != RecordTypeIdentifier.DATAEND
                   && nextRecord.RecordType != RecordTypeIdentifier.ROWHDR)
            {
                switch (nextRecord.RecordType)
                {
                    case RecordTypeIdentifier.CELLBLANK:
                    case RecordTypeIdentifier.CELLRK:
                    case RecordTypeIdentifier.CELLERROR:
                    case RecordTypeIdentifier.CELLBOOL:
                    case RecordTypeIdentifier.CELLREAL:
                    case RecordTypeIdentifier.CELLST:
                    case RecordTypeIdentifier.CELLISST:
                    case RecordTypeIdentifier.CELLFMLASTRING:
                    case RecordTypeIdentifier.CELLFMLANUM:
                    case RecordTypeIdentifier.CELLFMLABOOL:
                    case RecordTypeIdentifier.CELLFMLAERROR:
                        {
                            XlsbCell? cell = XlsbCell.ConstructCell(nextRecord, _instanceContext!);
                            int offset = cell?.ExcelColumnOffset ?? -1;
                            if (offset > 0 && offset <= _maxExcelColumnDimension)
                            {
                                localCells[offset] = cell;
                            }
                            else
                            {
                                // If the parsed cell offset is outside expected width, skip storing it to avoid OOB.
                                // (Should be rare, defensive programming.)
                            }
                        }
                        break;
                    default:
                        break;
                }
                nextRecord = _reader.ReadNextRecord();
            }
        }
        finally
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.ROWHDR)
            {
                _reader.RollBackLastRecord(nextRecord);
            }
            else
            {
                nextRecord.Dispose();
            }
        }
        // publish parsed cells once fully read to avoid partial-visible state
        _cells = localCells;
        _cellsLoaded = true;
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<ICell?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        if (_cells == null)
        {
            yield break;
        }

        for (int i = 1; i <= _maxExcelColumnDimension; i++)
        {
            yield return _cells[i];
        }
    }

    /// <InheritDoc />
    public IEnumerable<ICell?> GetAllCells(CancellationToken ct = default)
    {
        GetCells(ct);
        if (_cells == null)
        {
            yield break;
        }

        for (int i = 1; i <= _maxExcelColumnDimension; i++)
        {
            yield return _cells[i];
        }
    }

    /// <InheritDoc />
    public async Task<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        if (_cells == null || excelColumnIndex < 1 || excelColumnIndex > _maxExcelColumnDimension)
        {
            return null;
        }

        return _cells[excelColumnIndex];
    }

    /// <InheritDoc />
    public ICell? GetCell(int excelColumnIndex, CancellationToken ct = default)
    {
        GetCells(ct);
        if (_cells == null || excelColumnIndex < 1 || excelColumnIndex > _maxExcelColumnDimension)
        {
            return null;
        }

        return _cells[excelColumnIndex];
    }

    /// <InheritDoc />
    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default) => throw new NotImplementedException();

    /// <InheritDoc />
    public ICell? GetCell(string columnLetters, CancellationToken ct = default) => throw new NotImplementedException();
}