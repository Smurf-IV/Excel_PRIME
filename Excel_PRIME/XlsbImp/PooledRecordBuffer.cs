using System;
using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics;
using System.Runtime.CompilerServices;

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

    public ReadOnlySpan<byte> AsSpan(int offset, int length) => _array.AsSpan(offset, length);

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

    /// <summary>
    /// Get 32-bit integer from buffer at offset.
    /// Heavily inlined for performance.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public int GetInt32(int offset) => BinaryPrimitives.ReadInt32LittleEndian(_array.AsSpan(offset));

    /// <summary>
    /// Get 64-bit floating-point value from buffer at offset.
    /// Heavily inlined for performance.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public double GetDouble(int offset) => BinaryPrimitives.ReadDoubleLittleEndian(_array.AsSpan(offset));

    /// <summary>
    /// Get 16-bit integer from buffer at offset.
    /// Heavily inlined for performance.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public short GetInt16(int offset) => BinaryPrimitives.ReadInt16LittleEndian(_array.AsSpan(offset));

    /// <summary>
    /// Get single byte from buffer at offset.
    /// Heavily inlined for performance - single array access.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public byte GetByte(int offset) => _array[offset];

    /// <summary>
    /// Get string from buffer with UTF-16 LE encoding.
    /// Uses SIMD-optimized fast path for ASCII-only strings.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public string GetString(int offset)
    {
        int len = BinaryPrimitives.ReadInt32LittleEndian(_array.AsSpan(offset));
        // Use hybrid SIMD decoder - automatically optimizes for string length
        return SimdStringDecoder.DecodeUtf16WithHybridFastPath(_array, offset + 4, len * 2, len);
    }

    /// <summary>
    /// Get string from buffer with UTF-16 LE encoding and return the end offset.
    /// Uses SIMD-optimized fast path for ASCII-only strings.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public string? GetString(int offset, out int end)
    {
        int len = BinaryPrimitives.ReadInt32LittleEndian(_array.AsSpan(offset));
        if (len == -1)
        {
            end = offset + 4;
            return null;
        }
        end = offset + 4 + len * 2;
        // Use hybrid SIMD decoder - automatically optimizes for string length
        return SimdStringDecoder.DecodeUtf16WithHybridFastPath(_array, offset + 4, len * 2, len);
    }
}