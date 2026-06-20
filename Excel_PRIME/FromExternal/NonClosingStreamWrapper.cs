using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME.FromExternal;

/// <summary>
/// A wrapper for a <see cref="System.IO.Stream"/> that prevents the inner stream from being closed or disposed when the wrapper is disposed.
/// </summary>

public sealed class NonClosingStream : Stream
{
    private readonly Stream _inner;
    private bool _disposed;

    /// <inheritdoc/>
    public NonClosingStream(Stream inner)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
    }

    /// <inheritdoc/>
    public Stream InnerStream => _inner;

    /// <inheritdoc/>
    public override bool CanRead => !_disposed && _inner.CanRead;
    /// <inheritdoc/>
    public override bool CanSeek => !_disposed && _inner.CanSeek;
    /// <inheritdoc/>
    public override bool CanWrite => !_disposed && _inner.CanWrite;
    /// <inheritdoc/>
    public override long Length => _inner.Length;

    /// <inheritdoc/>
    public override long Position
    {
        get => _inner.Position;
        set => _inner.Position = value;
    }

    /// <inheritdoc/>
    public override void Flush() => _inner.Flush();

    /// <inheritdoc/>
    public override Task FlushAsync(CancellationToken cancellationToken) =>
        _inner.FlushAsync(cancellationToken);

    /// <inheritdoc/>
    public override int Read(byte[] buffer, int offset, int count) =>
        _inner.Read(buffer, offset, count);

    /// <inheritdoc/>
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
        _inner.ReadAsync(buffer, offset, count, cancellationToken);

    /// <inheritdoc/>
    public override long Seek(long offset, SeekOrigin origin) =>
        _inner.Seek(offset, origin);

    /// <inheritdoc/>
    public override void SetLength(long value) => _inner.SetLength(value);

    /// <inheritdoc/>
    public override void Write(byte[] buffer, int offset, int count) =>
        _inner.Write(buffer, offset, count);

    /// <inheritdoc/>
    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
        _inner.WriteAsync(buffer, offset, count, cancellationToken);

    /// <inheritdoc/>
    public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state) =>
        _inner.BeginRead(buffer, offset, count, callback, state);

    /// <inheritdoc/>
    public override int EndRead(IAsyncResult asyncResult) => _inner.EndRead(asyncResult);

    /// <inheritdoc/>
    public override IAsyncResult BeginWrite(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state) =>
        _inner.BeginWrite(buffer, offset, count, callback, state);

    /// <inheritdoc/>
    public override void EndWrite(IAsyncResult asyncResult) => _inner.EndWrite(asyncResult);

    /// <inheritdoc/>
#pragma warning disable CA2215 // The whole reason for doing this !
    protected override void Dispose(bool disposing)
#pragma warning restore CA2215
    {
        // Intentionally do not dispose the inner stream.
        // Optionally flush the inner stream if disposing is true.
        if (disposing && !_disposed)
        {
            try
            {
                _inner.Flush();
            }
            catch
            {
                // swallow flush exceptions
            }
        }

        _disposed = true;
        // Do not call base.Dispose to avoid closing inner stream.
    }

    /// If you want to allow explicit close of the inner stream:
    public void CloseInnerStream()
    {
        _disposed = true;
        _inner.Close();
    }

    /// <inheritdoc/>
    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) =>
        _inner.ReadAsync(buffer, cancellationToken);

    /// <inheritdoc/>
    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) =>
        _inner.WriteAsync(buffer, cancellationToken);
}
