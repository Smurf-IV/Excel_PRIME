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
    private Dictionary<int, Cell> _cells = [];
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
                _cells = null!;
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

    /// <InheritDoc />
    public async IAsyncEnumerable<ICell?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        for (int i = 1; i <= _maxExcelColumnDimension; i++)
        {
            _cells.TryGetValue(i, out Cell? found);
            yield return found;
        }
    }

    public IEnumerable<ICell?> GetAllCells(CancellationToken ct = default)
    {
        GetCellsAsync(ct).GetAwaiter().GetResult();
        for (int i = 1; i <= _maxExcelColumnDimension; i++)
        {
            _cells.TryGetValue(i, out Cell? found);
            yield return found;
        }
    }

    private const int bufferSize = 64;

    internal async Task GetCellsAsync(CancellationToken ct)
    {
        if (_reader.IsEmptyElement)
        {
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
        char[] buffer = ArrayPool<char>.Shared.Rent(bufferSize);
        StringBuilder valueBuilder = new();

        try
        {
            while (await _reader.ReadAsync().ConfigureAwait(false)
                       && !ct.IsCancellationRequested
                    && _reader.Depth > currentDepth
                      )
            {
                if (_reader.NodeType == XmlNodeType.Element
                    && ReferenceEquals(_reader.LocalName, _readerAtoms.cRefAtom)
                    && !_reader.IsEmptyElement  // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    )
                {
                    Cell? cell = await Cell.ConstructCellAsync(_reader, _instanceContext, _readerAtoms, buffer, valueBuilder).ConfigureAwait(false);
                    if (cell != null) // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    {
                        _cells.Add(cell.ExcelColumnOffset, cell);
                    }

                    valueBuilder.Length = 0;
                }
            }
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer);
        }
    }

    internal void GetCells(CancellationToken ct)
    {
        if (_reader.IsEmptyElement)
        {
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
        char[] buffer = ArrayPool<char>.Shared.Rent(bufferSize);
        StringBuilder valueBuilder = new();

        try
        {
            while (_reader.Read()
                   && !ct.IsCancellationRequested
                   && _reader.Depth > currentDepth
                  )
            {
                if (_reader.NodeType == XmlNodeType.Element
                    && ReferenceEquals(_reader.LocalName, _readerAtoms.cRefAtom)
                    && !_reader.IsEmptyElement  // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                   )
                {
                    Cell? cell = Cell.ConstructCell(_reader, _instanceContext, _readerAtoms, buffer, valueBuilder);
                    if (cell != null) // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                    {
                        _cells.Add(cell.ExcelColumnOffset, cell);
                    }

                    valueBuilder.Length = 0;
                }
            }
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer);
        }
    }

    /// <InheritDoc />
    public async Task<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        _cells.TryGetValue(excelColumnIndex, out Cell? found);
        return found;
    }

    /// <InheritDoc />
    public ICell? GetCell(int excelColumnIndex, CancellationToken ct = default)
    {
        GetCells(ct);
        _cells.TryGetValue(excelColumnIndex, out Cell? found);
        return found;
    }

    /// <InheritDoc />
    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default) => throw new NotImplementedException();

    /// <InheritDoc />
    public ICell? GetCell(string columnLetters, CancellationToken ct = default) => throw new NotImplementedException();
}
