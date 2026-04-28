using System;
using System.Buffers;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Implementation;

internal sealed class Row : IRowAsync
{
    [ThreadStatic]
    private static Row? t_row;

    private XmlReader? _reader;
    private InstanceContext _instanceContext = null!;
    private int _maxExcelColumnDimension;
    private bool _isDisposed;
    private Cell?[]? _cells;
    private bool _cellsLoaded;
    private ReaderAtoms _readerAtomsRefForSafety;

    private Row()
    {
        // Private ctor for pooling. Keep lightweight.
    }

    // CHANGED: Removed AggressiveOptimization - simple method that should inline naturally
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static Row Rent()
    {
        Row? sb = t_row;
        if (sb == null)
        {
            return new Row();
        }

        t_row = null;
        return sb;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void Return(Row? row)
    {
        if (row == null)
        {
            return;
        }

        row.Reset();
        // Replace any existing thread-local row (drop the previous one).
        t_row = row;
    }

    // CHANGED: Removed AggressiveOptimization - let JIT make optimal decisions for this initialization method
    internal void Initialize(XmlReader rowElement, InstanceContext instanceContext, int maxColumnDimension, ReaderAtoms readerAtoms)
    {
        _reader = rowElement;
        _instanceContext = instanceContext;
        _maxExcelColumnDimension = maxColumnDimension;
        // keep a ref only to help with defensive checks; main use is in comparisons inside methods
        // (ReaderAtoms itself holds references to the atomized names).
        _readerAtomsRefForSafety = readerAtoms;

        // Read initial attributes (previously in ctor)
        if (_reader.NodeType == XmlNodeType.Element
            && ReferenceEquals(_reader.LocalName, readerAtoms.rowRefAtom)
           )
        {
            int expectedAttributes = 1;
            while (_reader.MoveToNextAttribute() && expectedAttributes > 0)
            {
                // Retrieve the atomized name directly.
                string currentAttributeName = _reader.LocalName;
                if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
                {
                    RowOffset = _reader.Value.IntParse();
                    expectedAttributes--;
                }
                else if (ReferenceEquals(currentAttributeName, readerAtoms.hiddenRefAtom))
                {
                    // TODO: Do something about this
                    //_isCurrentRowHidden = ReadBooleanValue(_reader, buffer);
                }
            }

            _reader.MoveToElement();
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void Reset()
    {
        _isDisposed = false;
        _reader = null;
        _instanceContext = null!;
        _maxExcelColumnDimension = 0;
        _cells = null;
        _cellsLoaded = false;
        RowOffset = 0;
    }

    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _isDisposed = true;
        // Return to pool for reuse
        Return(this);
    }

    /// <InheritDoc />
    public int RowOffset { get; private set; }

    private const int BufferSize = 512;

    /// <summary>
    /// Ensure cells are read once. Cells are stored in a small array indexed by excel 1-based column offset.
    /// Using an array avoids Dictionary overhead and reduces per-row allocations for typical sheet widths.
    /// </summary>
    // CHANGED: Removed AggressiveOptimization - async state machine benefits from default JIT optimization
    internal async ValueTask GetCellsAsync(CancellationToken ct)
    {
        if (_cellsLoaded)
        {
            return;
        }

        if (_reader == null)
        {
            return;
        }

        if (_reader.IsEmptyElement)
        {
            _cells = new Cell?[_maxExcelColumnDimension];
            _cellsLoaded = true;
            return;
        }

        int currentDepth = _reader.Depth;
        if (_reader.NodeType != XmlNodeType.Element)
        {
            if (_reader.ReadState != 0)
            {
                return;
            }

            currentDepth--;
        }

        // Defer allocating the cell array until we actually parse cells to keep Row light-weight when unused.
        Cell?[] localCells = new Cell?[_maxExcelColumnDimension];
        char[] buffer = ArrayPool<char>.Shared.Rent(BufferSize);
        StringBuilder valueBuilder = ThreadStringBuilderPool.Rent();

        try
        {
            while (await _reader.ReadAsync().ConfigureAwait(false)
                       && !ct.IsCancellationRequested
                       && _reader.Depth > currentDepth)
            {
                if (_reader.NodeType == XmlNodeType.Element
                    && ReferenceEquals(_reader.LocalName, _readerAtomsRefForSafety.cRefAtom)
                    && !_reader.IsEmptyElement    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = await Cell.ConstructCellAsync(_reader, _instanceContext, _readerAtomsRefForSafety, buffer, valueBuilder).ConfigureAwait(false);
                    if (cell != null)    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    {
                        int offset = cell.ExcelColumnOffset - 1;
                        if (offset >= 0 && offset < _maxExcelColumnDimension)
                        {
                            localCells[offset] = cell;
                        }
                        else
                        {
                            // If the parsed cell offset is outside expected width, skip storing it to avoid OOB.
                            // (Should be rare, defensive programming.)
                        }
                    }
                }
            }

            if (_instanceContext.Options.ReturnDBNull)
            {
                for (int index = 0; index < localCells.Length; index++)
                {
                    localCells[index] ??= new Cell
                    {
                        CellValue = new CellValue(DBNull.Value, 0),
                        ExcelColumnOffset = index + 1
                    };
                }
            }
            // publish parsed cells once fully read to avoid partial-visible state
            _cells = localCells;
            _cellsLoaded = true;
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer);
            ThreadStringBuilderPool.Return(valueBuilder);
        }
    }

    // CHANGED: Removed AggressiveOptimization - let JIT optimize based on actual call patterns
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

        if (_reader.IsEmptyElement)
        {
            _cells = new Cell?[_maxExcelColumnDimension];
            _cellsLoaded = true;
            return;
        }

        int currentDepth = _reader.Depth;
        if (_reader.NodeType != XmlNodeType.Element)
        {
            if (_reader.ReadState != 0)
            {
                return;
            }

            currentDepth--;
        }

        Cell?[] localCells = new Cell?[_maxExcelColumnDimension];
        char[] buffer = ArrayPool<char>.Shared.Rent(BufferSize);
        StringBuilder valueBuilder = ThreadStringBuilderPool.Rent();

        try
        {
            while (_reader.Read()
                   && !ct.IsCancellationRequested
                   && _reader.Depth > currentDepth)
            {
                if (_reader.NodeType == XmlNodeType.Element
                    && ReferenceEquals(_reader.LocalName, _readerAtomsRefForSafety.cRefAtom)
                    && !_reader.IsEmptyElement  // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = Cell.ConstructCell(_reader, _instanceContext, _readerAtomsRefForSafety, buffer, valueBuilder);
                    if (cell != null) // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    {
                        int offset = cell.ExcelColumnOffset - 1;
                        if (offset >= 0 && offset < _maxExcelColumnDimension)
                        {
                            localCells[offset] = cell;
                        }
                    }
                }
            }

            if (_instanceContext.Options.ReturnDBNull)
            {
                for (int index = 0; index < localCells.Length; index++)
                {
                    localCells[index] ??= new Cell
                    {
                        CellValue = new CellValue(DBNull.Value, 0),
                        ExcelColumnOffset = index + 1
                    };
                }
            }
            _cells = localCells;
            _cellsLoaded = true;
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer);
            ThreadStringBuilderPool.Return(valueBuilder);
        }
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple wrapper that calls GetCellsAsync, should inline well
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ValueTask<IReadOnlyList<ICell?>?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_cellsLoaded)
        {
            return ValueTask.FromResult<IReadOnlyList<ICell?>?>(_cells);
        }

        return GetAllCellsAsyncCore(ct);
    }

    private async ValueTask<IReadOnlyList<ICell?>?> GetAllCellsAsyncCore(CancellationToken ct)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        return _cells;
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple wrapper that calls GetCells, should inline well
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public IReadOnlyList<ICell?>? GetAllCells(CancellationToken ct = default)
    {
        GetCells(ct);
        return _cells;
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple accessor with bounds check, inline is sufficient
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ValueTask<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        if (_cellsLoaded)
        {
            if (_cells == null || excelColumnIndex < 1 || excelColumnIndex > _maxExcelColumnDimension)
            {
                return ValueTask.FromResult<ICell?>(result: null);
            }
            return ValueTask.FromResult<ICell?>(_cells[excelColumnIndex - 1]);
        }

        return GetCellAsyncCore(excelColumnIndex, ct);
    }

    private async ValueTask<ICell?> GetCellAsyncCore(int excelColumnIndex, CancellationToken ct)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        if (_cells == null || excelColumnIndex < 1 || excelColumnIndex > _maxExcelColumnDimension)
        {
            return null;
        }
        return _cells[excelColumnIndex - 1];
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple accessor with bounds check, inline is sufficient
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

    // CHANGED: Removed AggressiveOptimization - tight loop benefits more from size reduction for i-cache
    public void CopyBoxedToArray(object?[] values, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(values);
        GetCells(ct);

        if (_cells == null)
        {
            throw new InvalidOperationException("Cells are not initialized.");
        }

        int minLength = Math.Min(values.Length, _maxExcelColumnDimension);
        for (int ordinal = 0; ordinal < minLength; ++ordinal)
        {
            values[ordinal] = _cells[ordinal]?.CellValue?.BoxedValue;
        }
    }
}