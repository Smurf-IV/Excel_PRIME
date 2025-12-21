using System;
using System.Buffers;
using System.IO;
using System.Runtime.CompilerServices;
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
    private PooledRecordBuffer? _rollBackRecord;
    
    // Reusable single-byte buffer to reduce allocations and improve cache locality
    private readonly byte[] _singleByteBuffer = new byte[1];

    /// <summary>
    /// Empty stream for unknown files
    /// </summary>
    public XlsbStreamReader()
    {
        _stream = null!;
    }

    /// <summary>
    /// Stream is NOT owned by this class and should be disposed by the caller.
    /// </summary>
    /// <param name="stream"></param>
    public XlsbStreamReader(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        // For modern hardware in 2025, 65536(64KB) is the standard "sweet spot" for many workloads
        _stream = new BufferedStream(stream, 64 * 1024);
    }

    /// <summary>
    /// Enable the return of the next RowHdr etc., without the need to re-read from the stream.
    /// </summary>
    /// <param name="record"></param>
    public void RollBackLastRecord(PooledRecordBuffer record) => _rollBackRecord = record;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private byte[] RentAndRead(int count)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        byte[] buffer = ArrayPool<byte>.Shared.Rent(count);
        int read = 0;
        while (read < count)
        {
            int bytesRead = _stream.Read(buffer, read, count - read);
            if (bytesRead == 0)
            {
                ArrayPool<byte>.Shared.Return(buffer);
                throw new EndOfStreamException();
            }
            read += bytesRead;
        }

        return buffer;
    }

    private async Task<byte[]> RentAndReadAsync(int count, CancellationToken ct)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(count);

        byte[] buffer = ArrayPool<byte>.Shared.Rent(count);
        int read = 0;
        while (read < count)
        {
            int bytesRead = await _stream.ReadAsync(buffer.AsMemory(read, count- read), ct).ConfigureAwait(false);
            if (bytesRead == 0)
            {
                ArrayPool<byte>.Shared.Return(buffer);
                throw new EndOfStreamException();
            }
            read += bytesRead;
        }

        return buffer;
    }

    public async Task<PooledRecordBuffer> ReadNextRecordAsync(CancellationToken ct)
    {
        if (_rollBackRecord != null)
        {
            PooledRecordBuffer record = _rollBackRecord;
            _rollBackRecord = null;
            return record;
        }
        if (!ReadRecordType(out RecordTypeIdentifier recordType)
            || recordType == RecordTypeIdentifier.EOF 
            || !ReadRecordLen(out uint recordLength)
            )
        {
            return new PooledRecordBuffer(recordType, succeeded: recordType != RecordTypeIdentifier.EOF);
        }

        try
        {
            byte[] buffer = await RentAndReadAsync((int)recordLength, ct).ConfigureAwait(false);
            return new PooledRecordBuffer(recordType, buffer, true);
        }
        catch
        {
            return new PooledRecordBuffer(RecordTypeIdentifier.EOF);
        }
    }

    public PooledRecordBuffer ReadNextRecord()
    {
        if (_rollBackRecord != null)
        {
            PooledRecordBuffer record = _rollBackRecord;
            _rollBackRecord = null;
            return record;
        }
        if (!ReadRecordType(out RecordTypeIdentifier recordType)
            || recordType == RecordTypeIdentifier.EOF
            || !ReadRecordLen(out uint recordLength)
           )
        {
            return new PooledRecordBuffer(recordType);
        }

        byte[] buffer = RentAndRead((int)recordLength);
        return new PooledRecordBuffer(recordType, buffer, true);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool ReadRecordType(out RecordTypeIdentifier recordType)
    {
        if (CarefulFieldRead( out uint value))
        {
            recordType = (RecordTypeIdentifier)(value);
            return true;
        }

        recordType = RecordTypeIdentifier.EOF;
        return false;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool ReadRecordLen(out uint recordLength) => CarefulFieldRead(out recordLength);

    /// <summary>
    /// Reads a variable-length encoded unsigned integer from the stream.
    /// Uses 7 bits per byte encoding with the MSB as a continuation flag.
    /// Optimized with reusable buffer to reduce allocations and improve cache locality.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private bool CarefulFieldRead(out uint value)
    {
        value = 0u;

        if (_stream.Read(_singleByteBuffer, 0, 1) == 0)
        {
            return false;
        }

        ref byte b1 = ref _singleByteBuffer[0];
        value = (uint)(b1 & 0x7F);

        if ((b1 & 0x80) == 0)
        {
            return true;
        }

        if (_stream.Read(_singleByteBuffer, 0, 1) == 0)
        {
            return false;
        }

        ref byte b2 = ref _singleByteBuffer[0];
        value = ((uint)(b2 & 0x7F) << 7) | value;

        if ((b2 & 0x80) == 0)
        {
            return true;
        }

        if (_stream.Read(_singleByteBuffer, 0, 1) == 0)
        {
            return false;
        }

        ref byte b3 = ref _singleByteBuffer[0];
        value = ((uint)(b3 & 0x7F) << 14) | value;

        if ((b3 & 0x80) == 0)
        {
            return true;
        }

        if (_stream.Read(_singleByteBuffer, 0, 1) == 0)
        {
            return false;
        }

        ref byte b4 = ref _singleByteBuffer[0];
        value = ((uint)(b4 & 0x7F) << 21) | value;

        return true;
    }
}