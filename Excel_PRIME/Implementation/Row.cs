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
    private readonly XmlReader _reader;
    private readonly InstanceContext _instanceContext;
    private readonly int _maxExcelColumnDimension;
    private bool _isDisposed;
    private Cell?[]? _cells;
    private bool _cellsLoaded;
    private readonly ReaderAtoms _readerAtoms;

    public Row(XmlReader rowElement, InstanceContext instanceContext, int maxColumnDimension, ReaderAtoms readerAtoms)
    {
        _reader = rowElement;
        _instanceContext = instanceContext;
        _maxExcelColumnDimension = maxColumnDimension;
        _readerAtoms = readerAtoms;

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

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                // Release references to allow GC of contained cells
                _cells = null;
                _cellsLoaded = false;
            }

            _isDisposed = true;
        }
    }

    ~Row()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }

    /// <InheritDoc />
    public int RowOffset { get; }

    private const int BufferSize = 64;

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
                    && ReferenceEquals(_reader.LocalName, _readerAtoms.cRefAtom)
                    && !_reader.IsEmptyElement  // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = await Cell.ConstructCellAsync(_reader, _instanceContext, _readerAtoms, buffer, valueBuilder).ConfigureAwait(false);
                    if (cell != null) // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
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
                    && ReferenceEquals(_reader.LocalName, _readerAtoms.cRefAtom)
                    && !_reader.IsEmptyElement  // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = Cell.ConstructCell(_reader, _instanceContext, _readerAtoms, buffer, valueBuilder);
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