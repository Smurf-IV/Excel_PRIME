using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

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
        _maxColumnDimension = maxColumnDimension;
        if (_reader is { NodeType: XmlNodeType.Element, LocalName: "row" })
        {
            while (_reader.MoveToNextAttribute())
            {
                switch (_reader.LocalName)
                {
                    case "r":
                        RowOffset = _reader.Value.IntParseUnsafe();
                        break;
                    case "hidden":
                        // TODO: Do something about this
                        //_isCurrentRowHidden = ReadBooleanValue(_reader, buffer);
                        break;
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

    private async Task GetCellsAsync(CancellationToken ct)
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
        while (await _reader.ReadAsync().ConfigureAwait(false)
               && !ct.IsCancellationRequested
                && _reader.Depth > currentDepth
               )
        {
            if (_reader is { NodeType: XmlNodeType.Element, LocalName: "c" })
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
    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }
}
