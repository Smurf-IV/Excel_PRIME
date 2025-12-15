using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

internal sealed class XlsbSheetReader : IOpenXmlSheetReaderAsync
{
    private readonly InstanceContext _instanceContext;
    private readonly XlsbStreamReader _reader;
    private bool _isDisposed;
    private readonly int _startRow;

    // Pool of Row instances shared by this reader (concurrent for safety).
    private readonly ConcurrentBag<XlsbRow> _rowPool = new();

    public XlsbSheetReader(Stream stream, InstanceContext instanceContext, CancellationToken ct)
    {
        _instanceContext = instanceContext;
        _reader = new XlsbStreamReader(stream);
        // Step into the worksheet
    }

    private XlsbRow CreateRowFromPool()
    {
        if (_rowPool.TryTake(out XlsbRow? r))
        {
            return r;
        }

        return XlsbRow.Rent();
    }

    private void ReturnRowToPool(XlsbRow r)
    {
        // Row.Dispose handles returning to global pool; but we keep an internal pool for speed.
        // Reset any reader-specific state is handled by Row.Reset inside Return.
        _rowPool.Add(r);
    }

    private bool ReadToNextStartRow(CancellationToken ct)
    {
        throw new NotImplementedException();
        return false;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _lastNullRow = null;    // Do not call dispose, because they have been returned to the caller
                _lastRow = null;    // Do not call dispose, because they have been returned to the caller
                // optionally clear local pool references so they can be GC'd
                while (_rowPool.TryTake(out _)) { }
            }

            _isDisposed = true;
        }
    }

    public (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// The Current row iterator offset (Starts at 1)
    /// </summary>
    public int CurrentRow { get; private set; }

    public void Dispose()
    {
        Dispose(isDisposing: true);
    }

#pragma warning disable CA2213  // Do not call dispose, because they are being returned to the caller
    private NullRow? _lastNullRow;
    private XlsbRow? _lastRow;
#pragma warning restore CA2213

    public async Task<IRowAsync?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        if (_lastRow != null)
        {
            CurrentRow++;
            if (_lastRow.RowOffset > CurrentRow)
            {
                _lastNullRow = new NullRow(CurrentRow);
                return _lastNullRow;
            }
            else
            {
                _lastNullRow = null;    // Do not call dispose, because they are being returned to the caller
                XlsbRow thisRow = _lastRow;
                _lastRow = null;    // Do not call dispose, because they are being returned to the caller
                return thisRow;
            }
        }

        if (CurrentRow < _startRow
            || !ReadToNextStartRow(ct)
           )
        {
            return null;
        }

        XlsbRow nextRow = CreateRowFromPool();
        nextRow.Initialize(_reader, _instanceContext, SheetDimensions.Width);

        if (cellGetMode > RowCellGet.None)
        {
            await nextRow.GetCellsAsync(ct).ConfigureAwait(false);
        }

        if (nextRow.RowOffset > CurrentRow)
        {
            // Deal with blank rows in the sheet?
            // i.e. ones that do not have a definition in the xml! Therefore, will "Look like a jump"
            _lastRow = nextRow;
            _lastNullRow = new NullRow(CurrentRow);
            return _lastNullRow;
        }

        return nextRow;
    }

    public IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        if (_lastRow != null)
        {
            CurrentRow++;
            if (_lastRow.RowOffset > CurrentRow)
            {
                _lastNullRow = new NullRow(CurrentRow);
                return _lastNullRow;
            }
            else
            {
                _lastNullRow = null;    // Do not call dispose, because they are being returned to the caller
                XlsbRow thisRow = _lastRow;
                _lastRow = null;    // Do not call dispose, because they are being returned to the caller
                return thisRow;
            }
        }

        if (CurrentRow < _startRow
            || !ReadToNextStartRow(ct)
           )
        {
            return null;
        }

        XlsbRow nextRow = CreateRowFromPool();
        nextRow.Initialize(_reader, _instanceContext, SheetDimensions.Width);

        if (cellGetMode > RowCellGet.None)
        {
            nextRow.GetCells(ct);
        }

        if (nextRow.RowOffset > CurrentRow)
        {
            // Deal with blank rows in the sheet?
            // i.e. ones that do not have a definition in the xml! Therefore, will "Look like a jump"
            _lastRow = nextRow;
            _lastNullRow = new NullRow(CurrentRow);
            return _lastNullRow;
        }

        return nextRow;
    }
}