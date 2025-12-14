using System;
using System.Buffers;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace ExcelPRIME.XlsbImp;

/// <summary>
/// Provides functionality for reading binary data from a stream in a structured manner.
/// </summary>
/// <remarks>
/// This class supports both synchronous and asynchronous methods for reading various data types,
/// including integers, floating-point numbers, and strings. The stream used by this class is not
/// owned by it and must be managed and disposed of by the caller.
/// </remarks>
/// <exception cref="ArgumentOutOfRangeException"></exception>
/// <exception cref="EndOfStreamException"></exception>
internal class XlsbStreamReader
{
    private readonly Stream _stream;
    private readonly Encoding _encoding = Encoding.Unicode; // use little endian byte order

    /// <summary>
    /// Stream is NOT owned by this class and should be disposed by the caller.
    /// </summary>
    /// <param name="stream"></param>
    public XlsbStreamReader(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        _stream = stream;
    }
    
    // Synchronous Methods
    public short ReadInt16()
    {
        byte[] buffer = RentAndRead(2);
        try
        {
            return BitConverter.ToInt16(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public ushort ReadUInt16()
    {
        byte[] buffer = RentAndRead(2);
        try
        {
            return BitConverter.ToUInt16(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public int ReadInt32()
    {
        byte[] buffer = RentAndRead(4);
        try
        {
            return BitConverter.ToInt32(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public uint ReadUInt32()
    {
        byte[] buffer = RentAndRead(4);
        try
        {
            return BitConverter.ToUInt32(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public long ReadInt64()
    {
        byte[] buffer = RentAndRead(8);
        try
        {
            return BitConverter.ToInt64(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public ulong ReadUInt64()
    {
        byte[] buffer = RentAndRead(8);
        try
        {
            return BitConverter.ToUInt64(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public float ReadSingle()
    {
        byte[] buffer = RentAndRead(4);
        try
        {
            return BitConverter.ToSingle(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public double ReadDouble()
    {
        byte[] buffer = RentAndRead(8);
        try
        {
            return BitConverter.ToDouble(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public string ReadString()
    {
        var length = ReadInt32()*2;

        ArgumentOutOfRangeException.ThrowIfNegative(length);

        byte[] buffer = RentAndRead(length);
        try
        {
            return _encoding.GetString(buffer, 0, length);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    
    private byte[] RentAndRead(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        byte[] buffer = ArrayPool<byte>.Shared.Rent(count);
        int bytesRead = _stream.Read(buffer, 0, count);
        if (bytesRead < count)
        {
            ArrayPool<byte>.Shared.Return(buffer);
            throw new EndOfStreamException();
        }
        return buffer;
    }
    // Asynchronous Methods
    public async Task<short> ReadInt16Async(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(2, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToInt16(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<ushort> ReadUInt16Async(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(2, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToUInt16(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<int> ReadInt32Async(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(4, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToInt32(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<uint> ReadUInt32Async(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(4, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToUInt32(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<long> ReadInt64Async(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(8, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToInt64(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<ulong> ReadUInt64Async(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(8, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToUInt64(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<float> ReadSingleAsync(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(4, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToSingle(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<double> ReadDoubleAsync(CancellationToken ct)
    {
        byte[] buffer = await RentAndReadAsync(8, ct).ConfigureAwait(false);
        try
        {
            return BitConverter.ToDouble(buffer, 0);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    public async Task<string> ReadStringAsync( CancellationToken ct)
    {
        int length = await ReadInt32Async(ct).ConfigureAwait(false) * 2;
        ArgumentOutOfRangeException.ThrowIfNegative(length);

        byte[] buffer = await RentAndReadAsync(length, ct).ConfigureAwait(false);
        try
        {
            return _encoding.GetString(buffer, 0, length);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
    private async Task<byte[]> RentAndReadAsync(int count, CancellationToken ct)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        byte[] buffer = ArrayPool<byte>.Shared.Rent(count);
        int bytesRead = await _stream.ReadAsync(buffer.AsMemory(0, count), ct).ConfigureAwait(false);
        if (bytesRead < count)
        {
            ArrayPool<byte>.Shared.Return(buffer);
            throw new EndOfStreamException();
        }
        return buffer;
    }
}

