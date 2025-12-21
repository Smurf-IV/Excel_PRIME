using System;
using System.Buffers;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelPRIME.XlsbImp;

[DebuggerDisplay("RecordType {RecordType}")]
internal sealed class PooledRecordBuffer : IDisposable
{
    private readonly byte[] _array;
    private bool _isDisposed;
    
    public PooledRecordBuffer(RecordTypeIdentifier recordType, byte[]? array = null, bool succeeded = false)
    {
        RecordType = recordType;
        _array = array!;
        Succeeded = succeeded;
    }

    public ref readonly byte this[int index] => ref _array[index];

    public RecordTypeIdentifier RecordType { get; }

    public bool Succeeded { get; }

    // Return the array to the ArrayPool
    public void Dispose()
    {
        if (!_isDisposed)
        {
            _isDisposed = true;
            if (_array != null!)
            {
                ArrayPool<byte>.Shared.Return(_array);
            }
        }
    }

    public int GetInt32(int offset) => BitConverter.ToInt32(_array, offset);

    public double GetDouble(int offset) => BitConverter.ToDouble(_array, offset);

    public short GetInt16(int offset) => BitConverter.ToInt16(_array, offset);

    public byte GetByte(int offset) => _array[offset];

    /// <summary>
    /// Get string from buffer with UTF-16 LE encoding.
    /// Uses optimized fast path for ASCII-only strings.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public string GetString(int offset)
    {
        int len = BitConverter.ToInt32(_array, offset);
        return DecodeUtf16WithFastPath(_array, offset + 4, len * 2, len);
    }

    /// <summary>
    /// Get string from buffer with UTF-16 LE encoding and return the end offset.
    /// Uses optimized fast path for ASCII-only strings.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public string? GetString(int offset, out int end)
    {
        int len = BitConverter.ToInt32(_array, offset);
        if (len == -1)
        {
            end = offset + 4;
            return null;
        }
        end = offset + 4 + len * 2;
        return DecodeUtf16WithFastPath(_array, offset + 4, len * 2, len);
    }

    /// <summary>
    /// Decode UTF-16 LE string with fast path for ASCII-only strings.
    /// 
    /// For ASCII-only strings (all characters in range 0x00-0x7F), the high byte
    /// of each UTF-16 LE code unit is zero. We detect this and use a faster
    /// decoding path that avoids the overhead of the full UTF-16 decoder.
    /// 
    /// This optimization is particularly valuable for XLSB files where most
    /// shared strings and column names are ASCII.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static string DecodeUtf16WithFastPath(byte[] buffer, int offset, int byteCount, int charCount)
    {
        ReadOnlySpan<byte> span = new ReadOnlySpan<byte>(buffer, offset, byteCount);

        // View bytes as UTF-16 LE code units (ushort)
        // This allows us to check the high byte of each character efficiently
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);

        // Fast ASCII detection: check if high byte of each UTF-16 LE code unit is zero
        // In UTF-16 LE, ASCII characters (0x00-0x7F) are stored as [low_byte, 0x00]
        for (int i = 0; i < units.Length; i++)
        {
            if ((units[i] & 0xFF00) != 0)
            {
                // Non-ASCII detected, fallback to standard UTF-16 decoder
                return Encoding.Unicode.GetString(buffer, offset, byteCount);
            }
        }

        // All characters are ASCII-only. 
        // Cast UTF-16 bytes directly to char span and create string.
        // This avoids the overhead of the full UTF-16 decoder.
        ReadOnlySpan<char> chars = MemoryMarshal.Cast<byte, char>(span);
        return new string(chars.Slice(0, charCount));
    }
}