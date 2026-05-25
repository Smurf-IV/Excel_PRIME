using System.Buffers;
using System.IO;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.Implementation;

internal sealed class ZipReaderAsync : IZipReaderAsync
{
    private bool _isDisposed;
    private ZipArchive? _archive;
    private const int BufferSize = 81920 / 2; // Stop internal ArrayPool from doubling and bursting into the LOH

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void OpenArchive(Stream fileStream, CancellationToken ct) =>
        _archive = new ZipArchive(fileStream, ZipArchiveMode.Read, true);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Task OpenArchiveAsync(Stream fileStream, CancellationToken ct) =>
        Task.Run(() => _archive = new ZipArchive(fileStream, ZipArchiveMode.Read, true),
            ct);

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public bool CopyTo(string entryName, Stream targetStream, CancellationToken ct)
    {
        ZipArchiveEntry? entry = _archive!.GetEntry(entryName);
        if (entry == null)
        {
            return false;
        }
        targetStream.SetLength(entry.Length);   // Just saves a few moments.. And also checks if the OS can actually take the size
        using Stream decompressor = entry.Open();
        //decompressor.CopyTo(targetStream, 81920 / 2);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(BufferSize);
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

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
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
        await decompressor.CopyToAsync(targetStream, BufferSize, ct)
                .ConfigureAwait(false);
        await decompressor.FlushAsync(ct).ConfigureAwait(false);
        return true;
    }

    // One day it will be async in the zipReader
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Task<Stream?> GetEntryAsync(string entryName, CancellationToken ct)
        => Task.FromResult(GetEntry(entryName));

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
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
