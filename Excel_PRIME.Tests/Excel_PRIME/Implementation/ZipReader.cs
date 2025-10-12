using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.Implementation;

internal sealed class ZipReader : IZipReader
{
    private bool _isDisposed;
    private ZipArchive? _archive;

    public Task OpenArchiveAsync(Stream fileStream, CancellationToken ct)
    {
        return Task.Run(() => _archive = new ZipArchive(fileStream, ZipArchiveMode.Read, true),
            ct);
    }

    public async Task<bool> CopyToAsync(string entryName, Stream targetStream, CancellationToken ct)
    {
        ZipArchiveEntry? entry = _archive!.GetEntry(entryName);
        if (entry == null)
        {
            return false;
        }

        using Stream source = entry.Open();
        await source.CopyToAsync(targetStream, ct)
                .ConfigureAwait(false);
        await source.FlushAsync(ct).ConfigureAwait(false);
        return true;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _archive?.Dispose();
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            _isDisposed = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~InternalZipReader()
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
