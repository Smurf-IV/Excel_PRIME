using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

internal sealed class XlsbLazyLoadSharedStrings : ISharedString
{
    private static readonly SemaphoreLocker _locker = new();
    private readonly Stream? _stream;
    private readonly XlsbStreamReader _reader;
    private readonly List<string> _currentlyLoaded;
    private bool _isDisposed;
    private readonly StringBuilder _currentStNodeBuilder = new();

    public XlsbLazyLoadSharedStrings()
    {
        _currentlyLoaded = [];
        _stream = null;
        _reader = new XlsbStreamReader();
    }

    public XlsbLazyLoadSharedStrings(Stream stream, CancellationToken ct)
    {
        _stream = stream;
        _reader = new XlsbStreamReader(stream);
        // advance to the content
        using PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        if ( !nextRecord.Succeeded || nextRecord.RecordType != RecordTypeIdentifier.SSTBEGIN)
        {
            throw new InvalidDataException("The provided stream is not a valid XLSB shared strings stream.");
        }
        //int totalCount = nextRecord.GetInt32(0);
        int count = nextRecord.GetInt32(4);
        _currentlyLoaded = new List<string>(count);
    }

    // TODO: Should this be refactored to take a Cancellation Token
    public string? this[int requestIndex]
    {
        get
        {
            if (requestIndex < 0)
            {
                // TODO: Throw an exception ?
                return null;
            }

            // Many sheets may be attempting to get shared strings
            if (requestIndex >= _currentlyLoaded.Count)
            {
                _locker.Enter();
                try
                {
                    // Use additional offset to reduce locking intensity
                    LoadUntil(requestIndex+16);
                }
                finally
                {
                    _locker.Exit();
                }
            }

            if (requestIndex >= _currentlyLoaded.Count)
            {
                // TODO: Throw an exception ?
                return string.Empty;
            }
            else
            {
                return _currentlyLoaded[requestIndex];
            }
        }
    }

    // TODO: If passed the CancellationToken, should it also be Async ?
    private void LoadUntil(int untilIndex)
    {
        // Parse sequentially until we have loaded enough shared strings.
        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        while (untilIndex >= _currentlyLoaded.Count
               && nextRecord is { Succeeded: true, RecordType: RecordTypeIdentifier.STRINGITEM })
        {
            //var flags = nextRecord.GetByte(0);
            string str = nextRecord.GetString(1);

            _currentlyLoaded.Add(str);
            nextRecord = _reader.ReadNextRecord();
        }
        nextRecord.Dispose();
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

    public void Dispose() => Dispose(isDisposing: true);
}