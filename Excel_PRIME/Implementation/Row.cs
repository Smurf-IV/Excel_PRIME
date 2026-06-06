using System.Buffers;
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
    private Cell[]? _cells;
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
        if (_cells != null)
        {
            ArrayPool<Cell>.Shared.Return(_cells);
            _cells = null;
        }
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

    // Reduced per-row ReadValueChunk buffer size (down to 128) to lower rented char[] pressure.
    private const int BufferSize = 128;

    /// <summary>
    /// Ensure cells are read once. Cells are stored in a small array indexed by excel 1-based column offset.
    /// Using an array avoids Dictionary overhead and reduces per-row allocations for typical sheet widths.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
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
            _cells = ArrayPool<Cell>.Shared.Rent(_maxExcelColumnDimension);
            _cells.AsSpan(0, _maxExcelColumnDimension).Clear();
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
        Cell[] localCells = ArrayPool<Cell>.Shared.Rent(_maxExcelColumnDimension);
        localCells.AsSpan(0, _maxExcelColumnDimension).Clear();
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
                    Cell cell = await Cell.ConstructCellAsync(_reader, _instanceContext, _readerAtomsRefForSafety, buffer, valueBuilder).ConfigureAwait(false);
                    if (!cell.CellValue.IsUnknown)    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
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
                for (int index = 0; index < _maxExcelColumnDimension; index++)
                {
                    if (localCells[index].CellValue.IsUnknown)
                    {
                        localCells[index] = new Cell(CellValue.GetDBNull(0), index + 1, CellType.Unknown);
                    }
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

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
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
            _cells = ArrayPool<Cell>.Shared.Rent(_maxExcelColumnDimension);
            _cells.AsSpan(0, _maxExcelColumnDimension).Clear();
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

        Cell[] localCells = ArrayPool<Cell>.Shared.Rent(_maxExcelColumnDimension);
        localCells.AsSpan(0, _maxExcelColumnDimension).Clear();
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
                    Cell cell = Cell.ConstructCell(_reader, _instanceContext, _readerAtomsRefForSafety, buffer, valueBuilder);
                    if (!cell.CellValue.IsUnknown) // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
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
                for (int index = 0; index < _maxExcelColumnDimension; index++)
                {
                    if (localCells[index].CellValue.IsUnknown)
                    {
                        localCells[index] = new Cell(CellValue.GetDBNull(0), index + 1, CellType.Unknown);
                    }
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

    public ValueTask<ArraySegment<Cell>> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_cellsLoaded)
        {
            return ValueTask.FromResult(_cells == null ? default : new ArraySegment<Cell>(_cells, 0, _maxExcelColumnDimension));
        }

        return GetAllCellsAsyncCore(ct);
    }

    private async ValueTask<ArraySegment<Cell>> GetAllCellsAsyncCore(CancellationToken ct)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        return _cells == null ? default : new ArraySegment<Cell>(_cells, 0, _maxExcelColumnDimension);
    }

    /// <InheritDoc />
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public ArraySegment<Cell> GetAllCells(CancellationToken ct = default)
    {
        GetCells(ct);
        return _cells == null ? default : new ArraySegment<Cell>(_cells, 0, _maxExcelColumnDimension);
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple accessor with bounds check, inline is sufficient
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ValueTask<Cell> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        if (_cellsLoaded)
        {
            if (_cells == null || excelColumnIndex < 1 || excelColumnIndex > _maxExcelColumnDimension)
            {
                return ValueTask.FromResult<Cell>(default);
            }
            return ValueTask.FromResult<Cell>(_cells[excelColumnIndex - 1]);
        }

        return GetCellAsyncCore(excelColumnIndex, ct);
    }

    private async ValueTask<Cell> GetCellAsyncCore(int excelColumnIndex, CancellationToken ct)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        if (_cells == null || excelColumnIndex < 1 || excelColumnIndex > _maxExcelColumnDimension)
        {
            return default;
        }
        return _cells[excelColumnIndex - 1];
    }

    /// <InheritDoc />
    // CHANGED: Removed AggressiveOptimization - simple accessor with bounds check, inline is sufficient
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Cell GetCell(int excelColumnIndex, CancellationToken ct = default)
    {
        GetCells(ct);
        if (_cells == null
            || excelColumnIndex < 1
            || excelColumnIndex > _maxExcelColumnDimension)
        {
            return default;
        }

        return _cells[excelColumnIndex - 1];
    }

    /// <InheritDoc />
    /// <InheritDoc />
    public async ValueTask<Cell> GetCellAsync(string columnLetters, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(columnLetters);
        if (!_cellsLoaded)
        {
            await GetCellsAsync(ct).ConfigureAwait(false);
        }

        return await GetCellAsync(columnLetters.GetColNumber(), ct).ConfigureAwait(false);
    }

    /// <InheritDoc />
    public Cell GetCell(string columnLetters, CancellationToken ct = default)
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
            values[ordinal] = _cells[ordinal].CellValue.BoxedValue;
        }
    }
}