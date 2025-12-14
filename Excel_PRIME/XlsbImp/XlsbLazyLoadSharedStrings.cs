using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

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
        int count = 128;
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
                    // The "requestIndex >= _currentlyLoaded.Count" is also done internally, so no need to check again after locking
                    throw new NotImplementedException();
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

    public void Dispose() => Dispose(isDisposing: true);
}