using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using Excel_PRIME.Shared;

namespace Excel_PRIME.Implementation;

internal class Sheet : ISheet
{
    private bool _isDisposed;
    private readonly IXmlReader _xmlReader;
    private readonly TempFile _sourceFile;

    internal Sheet(TempFile sourceFile, IXmlReader xmlReader, string name)
    {
        _sourceFile = sourceFile;
        _xmlReader = xmlReader;
        Name = name;
    }

    public string Name { get; }

    public Task<object?> GetCellAsync(int rowIndex, int columnIndex, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    public IAsyncEnumerable<object?[]> GetDefinedRangeAsync(in string range, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    public Task<object?> GetRangeCellAsync(in string rnageCell, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    public IAsyncEnumerable<object?[]> GetRowDataAsync(int skipRows, int startColumn, int numberOfColumns, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    protected virtual void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _xmlReader.Dispose();
                _sourceFile.Dispose();
            }

            _isDisposed = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~Sheet()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }
}
