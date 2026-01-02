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
    private readonly ConcurrentBag<XlsbRow> _rowPool = [];

    public XlsbSheetReader(BufferedStream stream, InstanceContext instanceContext, CancellationToken ct)
    {
        _instanceContext = instanceContext;
        _reader = new XlsbStreamReader(stream);
        bool foundSheetData = false;
        // Step into the worksheet
        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested
               && !foundSheetData)
        {
            switch (nextRecord.RecordType)
            {
                case RecordTypeIdentifier.SHEETPR:
                    {
                        // Step over the preAmble
                        string codeName = nextRecord.GetString(19);
                    }
                    break;

                case RecordTypeIdentifier.SHEETDATABEGIN:
                    // All is good ;-)
                    foundSheetData = true;
                    break;
                case RecordTypeIdentifier.DIMENSION:
                    {
                        // Read dimensions
                        _startRow = nextRecord.GetInt32(0);
                        int lastRow = nextRecord.GetInt32(4);
                        int lastCol = nextRecord.GetInt32(12);
                        SheetDimensions = (lastRow + 1, lastCol + 1); // Make them VBA Excel references
                    }
                    break;
                case RecordTypeIdentifier.COLINFO:
                    // We can ignore column info for now
                    break;
            }

            if (!foundSheetData)
            {
                nextRecord = _reader.ReadNextRecord();
            }
        }
        nextRecord.Dispose();
        CurrentRow = 0;
    }

    private XlsbRow CreateRowFromPool() =>
        _rowPool.TryTake(out XlsbRow? r)
            ? r
            : XlsbRow.Rent();

    private void ReturnRowToPool(XlsbRow r) =>
        // Row.Dispose handles returning to global pool; but we keep an internal pool for speed.
        // Reset any reader-specific state is handled by Row.Reset inside Return.
        _rowPool.Add(r);

    private async Task<bool> ReadToNextStartRowAsync(CancellationToken ct)
    {
        PooledRecordBuffer nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        try
        {
            while (nextRecord.Succeeded
                   && !ct.IsCancellationRequested)
            {
                switch (nextRecord.RecordType)
                {
                    case RecordTypeIdentifier.ROWHDR:
                        _reader.RollBackLastRecord(nextRecord);
                        return true;
                    case RecordTypeIdentifier.DATAEND:
                        return false;
                }
                nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
            }
        }
        finally
        {
            CurrentRow++;   // No rows to read, or the Dimension is lying
            nextRecord.Dispose();
        }

        return false;
    }

    private bool ReadToNextStartRow(CancellationToken ct)
    {
        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        try
        {
            while (nextRecord.Succeeded
                   && !ct.IsCancellationRequested)
            {
                switch (nextRecord.RecordType)
                {
                    case RecordTypeIdentifier.ROWHDR:
                        _reader.RollBackLastRecord(nextRecord);
                        return true;
                    case RecordTypeIdentifier.DATAEND:
                        return false;
                }
                nextRecord = _reader.ReadNextRecord();
            }
        }
        finally
        {
            CurrentRow++;   // No rows to read, or the Dimension is lying
            nextRecord.Dispose();
        }

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
                while (_rowPool.TryTake(out _)) 
                {
                }
            }

            _isDisposed = true;
        }
    }

    public (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// The Current row iterator offset (Starts at 1)
    /// </summary>
    public int CurrentRow { get; private set; }

    public void Dispose() => Dispose(isDisposing: true);

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
            || !await ReadToNextStartRowAsync(ct).ConfigureAwait(false)
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