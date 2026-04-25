using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

internal sealed class XlsbRow : IRowAsync
{
    [ThreadStatic]
    private static XlsbRow? t_row;

    // Thread-local pool for cell arrays to reduce repeated allocations
    [ThreadStatic]
    private static XlsbCell?[]? t_cellArrayPool;

    private XlsbStreamReader? _reader;
    private InstanceContext? _instanceContext;
    private int _maxExcelColumnDimension;
    private bool _isDisposed;
    private XlsbCell?[]? _cells;
    private bool _cellsLoaded;

    private XlsbRow()
    {
        // Private ctor for pooling. Keep lightweight.
    }

    // CHANGED: Removed AggressiveOptimization - simple method that should inline naturally
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static XlsbRow Rent()
    {
        XlsbRow? sb = t_row;
        if (sb == null)
        {
            return new XlsbRow();
        }

        t_row = null;
        return sb;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Return(XlsbRow? row)
    {
        if (row == null)
        {
            return;
        }

        row.Reset();
        // Replace any existing thread-local row (drop the previous one).
        t_row = row;
    }

    // CHANGED: Removed implicit optimization - let JIT optimize initialization
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

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void Reset()
    {
        _isDisposed = false;

        // Return cell array to thread-local pool for reuse
        if (_cells != null)
        {
            t_cellArrayPool = _cells;
            _cells = null;
        }

        _reader = null;
        _instanceContext = null;
        _maxExcelColumnDimension = 0;
        _cellsLoaded = false;
        RowOffset = 0;
    }

    // CHANGED: Removed AggressiveOptimization - async state machine benefits from default JIT
    internal async Task GetCellsAsync(CancellationToken ct)
    {
        if (_cellsLoaded
            || _reader == null)
        {
            return;
        }

        // Reuse pooled array if available, otherwise allocate new one
        XlsbCell?[] localCells = t_cellArrayPool ?? new XlsbCell?[_maxExcelColumnDimension];

        // If pooled array was reused but is wrong size, resize it (rare case)
        if (localCells.Length != _maxExcelColumnDimension)
        {
            localCells = new XlsbCell?[_maxExcelColumnDimension];
        }
        else
        {
            // Clear pooled array for reuse
            Array.Clear(localCells, 0, localCells.Length);
        }

        PooledRecordBuffer nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        try
        {
            while (nextRecord.Succeeded
                   && !ct.IsCancellationRequested
                   && nextRecord.RecordType != RecordTypeIdentifier.DATAEND
                   && nextRecord.RecordType != RecordTypeIdentifier.ROWHDR)
            {
                // Fast path: skip non-cell records
                RecordTypeIdentifier recordType = nextRecord.RecordType;
                if (recordType is < RecordTypeIdentifier.CELLBLANK or > RecordTypeIdentifier.CELLFMLAERROR)
                {
                    nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
                    continue;
                }
                XlsbCell? cell = XlsbCell.ConstructCell(nextRecord, _instanceContext!);
                if (cell != null)
                {
                    int offset = cell.ExcelColumnOffset - 1;
                    if (offset >= 0 && offset < _maxExcelColumnDimension)
                    {
                        localCells[offset] = cell;
                    }
                    // Out-of-range cells are silently dropped (defensive programming)
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

        if (_instanceContext.Options.ReturnDBNull)
        {
            for (int index = 0; index < localCells.Length; index++)
            {
                localCells[index] ??= new XlsbCell
                {
                    CellValue = new CellValue(DBNull.Value, 0),
                    ExcelColumnOffset = index + 1
                };
            }
        }
        // Publish parsed cells once fully read to avoid partial-visible state
        _cells = localCells;
        t_cellArrayPool = null; // Mark pooled array as consumed
        _cellsLoaded = true;
    }

    // CHANGED: Removed AggressiveOptimization - let JIT optimize based on actual call patterns
    internal void GetCells(CancellationToken ct)
    {
        if (_cellsLoaded
            || _reader == null)
        {
            return;
        }

        // Reuse pooled array if available, otherwise allocate new one
        XlsbCell?[] localCells = t_cellArrayPool ?? new XlsbCell?[_maxExcelColumnDimension];

        // If pooled array was reused but is wrong size, resize it (rare case)
        if (localCells.Length != _maxExcelColumnDimension)
        {
            localCells = new XlsbCell?[_maxExcelColumnDimension];
        }
        else
        {
            // Clear pooled array for reuse
            Array.Clear(localCells, 0, localCells.Length);
        }

        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        try
        {
            while (nextRecord.Succeeded
                   && !ct.IsCancellationRequested
                   && nextRecord.RecordType != RecordTypeIdentifier.DATAEND
                   && nextRecord.RecordType != RecordTypeIdentifier.ROWHDR)
            {
                // Fast path: skip non-cell records
                RecordTypeIdentifier recordType = nextRecord.RecordType;
                if (recordType is < RecordTypeIdentifier.CELLBLANK or > RecordTypeIdentifier.CELLFMLAERROR)
                {
                    nextRecord = _reader.ReadNextRecord();
                    continue;
                }
                XlsbCell? cell = XlsbCell.ConstructCell(nextRecord, _instanceContext!);
                if (cell != null)
                {
                    int offset = cell.ExcelColumnOffset - 1;
                    if (offset >= 0 && offset < _maxExcelColumnDimension)
                    {
                        localCells[offset] = cell;
                    }
                    // Out-of-range cells are silently dropped (defensive programming)
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

        if (_instanceContext.Options.ReturnDBNull)
        {
            for (int index = 0; index < localCells.Length; index++)
            {
                localCells[index] ??= new XlsbCell
                {
                    CellValue = new CellValue(DBNull.Value, 0),
                    ExcelColumnOffset = index + 1
                };
            }
        }
        // Publish parsed cells once fully read to avoid partial-visible state
        _cells = localCells;
        t_cellArrayPool = null; // Mark pooled array as consumed
        _cellsLoaded = true;
    }

    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }
        _isDisposed = true;
        Return(this);
    }

    /// <InheritDoc />
    public int RowOffset { get; private set; }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple wrapper, inline better
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public async ValueTask<IReadOnlyList<ICell?>?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        return _cells;
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple wrapper, inline better
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IReadOnlyList<ICell?>? GetAllCells(CancellationToken ct = default)
    {
        GetCells(ct);
        return _cells;
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple accessor, inline better
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public async ValueTask<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        if (_cells == null
            || excelColumnIndex < 1
            || excelColumnIndex > _maxExcelColumnDimension)
        {
            return null;
        }

        return _cells[excelColumnIndex - 1];
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple accessor, inline better
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ICell? GetCell(int excelColumnIndex, CancellationToken ct = default)
    {
        GetCells(ct);
        if (_cells == null
            || excelColumnIndex < 1
            || excelColumnIndex > _maxExcelColumnDimension)
        {
            return null;
        }

        return _cells[excelColumnIndex - 1];
    }

    /// <InheritDoc />
    public async ValueTask<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(columnLetters);
        if (!_cellsLoaded)
        {
            await GetCellsAsync(ct).ConfigureAwait(false);
        }

        return await GetCellAsync(columnLetters.GetColNumber(), ct).ConfigureAwait(false);
    }

    /// <InheritDoc />
    public ICell? GetCell(string columnLetters, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(columnLetters);

        if (!_cellsLoaded)
        {
            GetCells(ct);
        }

        return GetCell(columnLetters.GetColNumber(), ct);
    }

    // CHANGED: Removed AggressiveOptimization - tight loop benefits from smaller code
    public void CopyBoxedToArray(object?[] values, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(values);
        GetCells(ct);

        if (_cells == null)
        {
            throw new InvalidOperationException("Cells are not initialized.");
        }

        int minLength = Math.Min(values.Length, _cells.Length);
        for (int ordinal = 0; ordinal < minLength; ++ordinal)
        {
            values[ordinal] = _cells[ordinal]?.CellValue.BoxedValue;
        }
    }
}