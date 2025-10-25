using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal sealed class Row : IRow
{
    private readonly XmlReader _reader;
    private readonly ISharedString _sharedStrings;
    private readonly int _maxColumnDimension;
    private bool _isDisposed;
    private Dictionary<int, Cell>? _cells;

    public Row(XmlReader rowElement, ISharedString sharedStrings, int maxColumnDimension)
    {
        _reader = rowElement;
        _sharedStrings = sharedStrings;
        string rowName = _reader.NameTable.Add("row");
        string rRef = _reader.NameTable.Add("r");
        string hiddenName = _reader.NameTable.Add("hidden");
        _maxColumnDimension = maxColumnDimension;
        if (_reader.NodeType == XmlNodeType.Element
            && Object.ReferenceEquals(_reader.LocalName, rowName)
            )
        {
            while (_reader.MoveToNextAttribute())
            {
                // Retrieve the atomized name directly.
                string currentAttributeName = _reader.LocalName;
                if (Object.ReferenceEquals(currentAttributeName, rRef))
                {
                    RowOffset = _reader.Value.IntParseUnsafe();
                }
                else if (Object.ReferenceEquals(currentAttributeName, hiddenName))
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
                _cells = null;
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
        for (int i = 0; i < _maxColumnDimension; i++)
        {
            _cells.TryGetValue(i, out Cell? found);
            yield return found;
        }
    }

    public IEnumerable<ICell?> GetAllCells(CancellationToken ct = default)
    {
        GetCellsAsync(ct).GetAwaiter().GetResult();
        for (int i = 0; i < _maxColumnDimension; i++)
        {
            _cells.TryGetValue(i, out Cell? found);
            yield return found;
        }
    }

    internal async Task GetCellsAsync(CancellationToken ct)
    {
        if (_cells != null)
        {
            return;
        }
        _cells = new Dictionary<int, Cell>();
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

        string cRef = _reader.NameTable.Add("c");
        while (await _reader.ReadAsync().ConfigureAwait(false)
               && !ct.IsCancellationRequested
                && _reader.Depth > currentDepth
               )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && Object.ReferenceEquals(_reader.LocalName, cRef)
                )
            {
                Cell cell = await Cell.ConstructCellAsync(_reader, _sharedStrings).ConfigureAwait(false);
                _cells.Add(cell.ExcelColumnOffset - 1, cell);
            }
        }
    }

    /// <InheritDoc />
    public async Task<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        _cells.TryGetValue(excelColumnIndex - 1, out Cell? found);
        return found;
    }

    /// <InheritDoc />
    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default) => throw new NotImplementedException();
}
