using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

internal sealed class XlsbWorkBookReader : IOpenXmlWorkBookReaderAsync
{
    private readonly Stream _stream;
    private readonly XlsbStreamReader _reader;
    private bool _isDisposed;

    public XlsbWorkBookReader(Stream? stream, CancellationToken _)
    {
        _stream = stream!;
        _reader = new XlsbStreamReader(_stream);
    }

    public async IAsyncEnumerable<KeyValuePair<string, int>> GetSheetNamesAsync(
        [EnumeratorCancellation] CancellationToken ct)
    {
        int relativeSheetId = 0;

        PooledRecordBuffer nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        bool foundSheets = false;
        while (nextRecord.Succeeded)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BUNDLESHEET)
            {
                foundSheets = true;
                string? rel = nextRecord.GetString(8, out int next);
                string name = nextRecord.GetString(next);
                if (rel == null)
                {
                    // no sheet rel means it is a macro.
                }
                else
                {
                    relativeSheetId++;
                    // `r:id` and `sheetId` are not to be trusted
                    yield return new KeyValuePair<string, int>(name, relativeSheetId);
                }
            }

            nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
            if (foundSheets
                && nextRecord.RecordType != RecordTypeIdentifier.BUNDLESHEET
                )
            {
                break;
            }
        }
    }

    public IEnumerable<KeyValuePair<string, int>> GetSheetNames([EnumeratorCancellation] CancellationToken ct)
    {
        throw new NotImplementedException();
    }

    public async Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(
        IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        throw new NotImplementedException();
    }

    public IReadOnlyDictionary<string, DefinedRange> GetDefinedRanges(IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        throw new NotImplementedException();
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _stream?.Dispose();
            }

            _isDisposed = true;
        }
    }

    ~XlsbWorkBookReader()
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
}
