using System;
using System.Buffers;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.Implementation;

internal sealed class ZipReaderAsync : IZipReaderAsync
{
    private bool _isDisposed;
    private ZipArchive? _archive;

    public void OpenArchive(Stream fileStream, CancellationToken ct) =>
        _archive = new ZipArchive(fileStream, ZipArchiveMode.Read, true);

    public Task OpenArchiveAsync(Stream fileStream, CancellationToken ct) =>
        Task.Run(() => _archive = new ZipArchive(fileStream, ZipArchiveMode.Read, true),
            ct);

    public bool CopyTo(string entryName, Stream targetStream, CancellationToken ct)
    {
        ZipArchiveEntry? entry = _archive!.GetEntry(entryName);
        if (entry == null)
        {
            return false;
        }
        targetStream.SetLength(entry.Length);   // Just saves a few moments.. And also checks if the OS can actually take the size
        using Stream decompressor = entry.Open();
        const int bufferSize = 81920 / 2; // Stop internal ArrayPool from doubling and bursting into the LOH !!
        //decompressor.CopyTo(targetStream, 81920 / 2);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(bufferSize);
        try
        {
            int bytesRead;
            while ((bytesRead = decompressor.Read(buffer, 0, buffer.Length)) != 0)
            {
                if (ct.IsCancellationRequested)
                {
                    return false;
                }
                targetStream.Write(buffer, 0, bytesRead);
            }
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
        return true;
    }

    public async Task<bool> CopyToAsync(string entryName, Stream targetStream, CancellationToken ct)
    {
        ZipArchiveEntry? entry = _archive!.GetEntry(entryName);
        if (entry == null)
        {
            return false;
        }
        targetStream.SetLength(entry.Length);   // Just saves a few moments.. And also checks if the OS can actually take the size
#if NET10_0_OR_GREATER
        using Stream decompressor = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
        using Stream decompressor = entry.Open();
#endif
        // "81920 / 2" -> Stop internal ArrayPool from doubling and bursting into the LOH !!
        await decompressor.CopyToAsync(targetStream, 81920 / 2, ct)
                .ConfigureAwait(false);
        await decompressor.FlushAsync(ct).ConfigureAwait(false);
        return true;
    }

    public Stream? GetEntry(string entryName)
    {
        ZipArchiveEntry? entry = _archive!.GetEntry(entryName);
        return entry?.Open();
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

    ~ZipReaderAsync()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: false);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }
}
