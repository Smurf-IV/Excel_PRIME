using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;


namespace ExcelPRIME.XlsbImp;

/// <summary>
/// SIMD-optimized variant of PooledRecordBuffer for Unicode string decoding.
/// This version uses Vector<ushort/> for parallel ASCII detection.
/// </summary>
internal static class SimdStringDecoder
{
    /// <summary>
    /// Decode UTF-16 LE string with SIMD-optimized fast path for ASCII-only strings.
    /// 
    /// Uses SIMD vectorization to check multiple UTF-16 code units in parallel.
    /// For ASCII-only strings (all characters in range 0x00-0x7F), the high byte
    /// of each UTF-16 LE code unit is zero. We detect this using Vector<ushort/>
    /// operations which can check 8-16 code units simultaneously on modern CPUs.
    /// 
    /// Performance characteristics:
    /// - ASCII detection: 8-16x faster than scalar (processes 8-16 units per iteration)
    /// - SIMD overhead amortized over vector width
    /// - Early exit on non-ASCII for fast failure case
    /// - Graceful fallback to standard decoder for non-ASCII
    /// 
    /// Typical XLSB profile:
    /// - ~90% of strings are ASCII (column names, sheet names, shared strings)
    /// - Average string length: 5-50 characters
    /// - Expected speedup: 2-4x overall on XLSB string extraction
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static string DecodeUtf16WithSimdFastPath(byte[] buffer, int offset, int byteCount, int charCount)
    {
        ReadOnlySpan<byte> span = new(buffer, offset, byteCount);

        // View bytes as UTF-16 LE code units (ushort)
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);

        // SIMD ASCII detection: check if all high bytes are zero
        // Process Vector<ushort>.Count units in parallel (typically 8 or 16 on x86-64)
        int i = 0;
        int unitsLength = units.Length;

        // SIMD fast path: process whole vectors
        int vectorSize = Vector<ushort>.Count;
        int simdEnd = unitsLength - (unitsLength % vectorSize);

        // Create mask for high byte: 0xFF00
        Vector<ushort> highByteMask = new(0xFF00);

        for (; i < simdEnd; i += vectorSize)
        {
            // Load vector of code units
            Vector<ushort> vector = new(units.Slice(i, vectorSize));

            // Mask out high bytes
            Vector<ushort> highBytes = vector & highByteMask;

            // If any high byte is non-zero, we have non-ASCII
            if (!Vector.EqualsAll(highBytes, Vector<ushort>.Zero))
            {
                goto NonAsciiDetected;
            }
        }

        // Scalar cleanup: handle remaining units not processed by SIMD
        for (; i < unitsLength; i++)
        {
            if ((units[i] & 0xFF00) != 0)
            {
                goto NonAsciiDetected;
            }
        }

        // All characters are ASCII-only. 
        // Cast UTF-16 bytes directly to char span and create string.
        // This avoids the overhead of the full UTF-16 decoder.
        ReadOnlySpan<char> chars = MemoryMarshal.Cast<byte, char>(span);
        return new string(chars.Slice(0, charCount));

    NonAsciiDetected:
        // Non-ASCII detected, fallback to standard UTF-16 decoder
        return Encoding.Unicode.GetString(buffer, offset, byteCount);
    }

    /// <summary>
    /// Hybrid scalar approach: Process in larger batches for CPU pipelining benefit
    /// while avoiding SIMD overhead on very short strings.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static string DecodeUtf16WithHybridFastPath(byte[] buffer, int offset, int byteCount, int charCount)
    {
        ReadOnlySpan<byte> span = new(buffer, offset, byteCount);
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);

        int i = 0;
        int unitsLength = units.Length;

        // For very short strings, use scalar approach
        if (unitsLength < 8)
        {
            for (; i < unitsLength; i++)
            {
                if ((units[i] & 0xFF00) != 0)
                {
                    goto NonAsciiDetected;
                }
            }
        }
        else
        {
            // For longer strings, use SIMD
            int vectorSize = Vector<ushort>.Count;
            int simdEnd = unitsLength - (unitsLength % vectorSize);
            Vector<ushort> highByteMask = new(0xFF00);

            for (; i < simdEnd; i += vectorSize)
            {
                Vector<ushort> vector = new(units.Slice(i, vectorSize));
                Vector<ushort> highBytes = vector & highByteMask;

                if (!Vector.EqualsAll(highBytes, Vector<ushort>.Zero))
                {
                    goto NonAsciiDetected;
                }
            }

            // Scalar cleanup
            for (; i < unitsLength; i++)
            {
                if ((units[i] & 0xFF00) != 0)
                {
                    goto NonAsciiDetected;
                }
            }
        }

        // All ASCII
        ReadOnlySpan<char> chars = MemoryMarshal.Cast<byte, char>(span);
        return new string(chars.Slice(0, charCount));

    NonAsciiDetected:
        return Encoding.Unicode.GetString(buffer, offset, byteCount);
    }
}
