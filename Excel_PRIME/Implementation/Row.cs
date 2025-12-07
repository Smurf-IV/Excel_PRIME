using System;
using System.Buffers;
using System.Collections.Concurrent;
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
    private XmlReader? _reader;
    private InstanceContext? _instanceContext;
    private int _maxExcelColumnDimension;
    private bool _isDisposed;
    private Cell?[]? _cells;
    private bool _cellsLoaded;
    private ReaderAtoms _readerAtomsRefForSafety;

    // Small object pool for Row instances to avoid allocating a new Row per XML row.
    private static readonly ConcurrentBag<Row> s_pool = new();

    private Row()
    {
        // Private ctor for pooling. Keep lightweight.
    }

    internal static Row Rent()
    {
        if (s_pool.TryTake(out Row? item))
        {
            return item;
        }

        return new Row();
    }

    internal static void Return(Row row)
    {
        // Reset state so next consumer sees a clean Row.
        row.Reset();
        s_pool.Add(row);
    }

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

    private const int BufferSize = 512;

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

        if (_reader.IsEmptyElement)
        {
            _cells = new Cell?[_maxExcelColumnDimension + 1];
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
        Cell?[] localCells = new Cell?[_maxExcelColumnDimension + 1];
        char[] buffer = ArrayPool<char>.Shared.Rent(BufferSize);
        StringBuilder valueBuilder = ThreadStringBuilderPool.Rent();

        try
        {
            while (await _reader.ReadAsync().ConfigureAwait(false)
                       && !ct.IsCancellationRequested
                       && _reader.Depth > currentDepth)
            {
                if (_reader.NodeType == XmlNodeType.Element
                    && ReferenceEquals(_reader.LocalName, _readerAtomsRefForSafety!.cRefAtom)
                    && !_reader.IsEmptyElement    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = await Cell.ConstructCellAsync(_reader, _instanceContext!, _readerAtomsRefForSafety!, buffer, valueBuilder).ConfigureAwait(false);
                    if (cell != null)    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    {
                        int offset = cell.ExcelColumnOffset;
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

                    valueBuilder.Length = 0;
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
            _cells = new Cell?[_maxExcelColumnDimension + 1];
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

        Cell?[] localCells = new Cell?[_maxExcelColumnDimension + 1];
        char[] buffer = ArrayPool<char>.Shared.Rent(BufferSize);
        StringBuilder valueBuilder = ThreadStringBuilderPool.Rent();

        try
        {
            while (_reader.Read()
                   && !ct.IsCancellationRequested
                   && _reader.Depth > currentDepth)
            {
                if (_reader.NodeType == XmlNodeType.Element
                    && ReferenceEquals(_reader.LocalName, _readerAtomsRefForSafety!.cRefAtom)
                    && !_reader.IsEmptyElement  // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = Cell.ConstructCell(_reader, _instanceContext!, _readerAtomsRefForSafety, buffer, valueBuilder);
                    if (cell != null) // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    {
                        int offset = cell.ExcelColumnOffset;
                        if (offset > 0 && offset <= _maxExcelColumnDimension)
                        {
                            localCells[offset] = cell;
                        }
                    }

                    valueBuilder.Length = 0;
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