using System;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace ExcelPRIME.Bench;
/// <summary>
/// Benchmark for UTF-16 string extraction in XLSB format.
/// Tests baseline Encoding.Unicode.GetString vs optimized ASCII-only fast path.
/// </summary>
[ExcludeFromCodeCoverage]
[SimpleJob(warmupCount: 3, iterationCount: 5)]
public class StringExtractionBenchmark
{
    private const int NumSamples = 1000;
    private byte[][] asciiStrings = null !;
    private byte[][] latinStrings = null !;
    private byte[][] mixedStrings = null !;
    private byte[][] unicodeStrings = null !;
    [GlobalSetup]
    public void Setup()
    {
        asciiStrings = new byte[NumSamples][];
        latinStrings = new byte[NumSamples][];
        mixedStrings = new byte[NumSamples][];
        unicodeStrings = new byte[NumSamples][];
        for (int i = 0; i < NumSamples; i++)
        {
            // ASCII-only strings (most common case in Excel)
            asciiStrings[i] = EncodeUtf16String($"Column_{i}");
            // Latin-1 Extended (accented characters)
            latinStrings[i] = EncodeUtf16String($"Café_{i}");
            // Mixed ASCII and some non-ASCII
            mixedStrings[i] = EncodeUtf16String(i % 2 == 0 ? $"Data_{i}" : $"Datos_{i}");
            // Full Unicode with emoji
            unicodeStrings[i] = EncodeUtf16String($"Test_{i}_🎉");
        }
    }

    /// <summary>
    /// Encodes string as UTF-16 LE with 4-byte character count prefix (XLSB format).
    /// Format: [4-byte char count (LE)][UTF-16 LE data]
    /// </summary>
    private static byte[] EncodeUtf16String(string text)
    {
        byte[] encoded = Encoding.Unicode.GetBytes(text);
        byte[] result = new byte[encoded.Length + 4];
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(0, 4), text.Length);
        encoded.CopyTo(result, 4);
        return result;
    }

    // ===== ASCII String Benchmarks =====
    //[Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string BaselineAsciiStrings()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = asciiStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = Encoding.Unicode.GetString(data, 4, len * 2);
        }

        return result;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string OptimizedAsciiStringsWithFastPath()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = asciiStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = DecodeUtf16WithFastPath(data, 4, len * 2, len);
        }

        return result;
    }

    // ===== Latin-1 String Benchmarks =====
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string BaselineLatinStrings()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = latinStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = Encoding.Unicode.GetString(data, 4, len * 2);
        }

        return result;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string OptimizedLatinStringsWithFastPath()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = latinStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = DecodeUtf16WithFastPath(data, 4, len * 2, len);
        }

        return result;
    }

    // ===== Mixed String Benchmarks =====
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string BaselineMixedStrings()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = mixedStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = Encoding.Unicode.GetString(data, 4, len * 2);
        }

        return result;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string OptimizedMixedStringsWithFastPath()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = mixedStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = DecodeUtf16WithFastPath(data, 4, len * 2, len);
        }

        return result;
    }

    // ===== Unicode String Benchmarks =====
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string BaselineUnicodeStrings()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = unicodeStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = Encoding.Unicode.GetString(data, 4, len * 2);
        }

        return result;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public string OptimizedUnicodeStringsWithFastPath()
    {
        string result = "";
        for (int i = 0; i < NumSamples; i++)
        {
            byte[] data = unicodeStrings[i];
            int len = BinaryPrimitives.ReadInt32LittleEndian(data.AsSpan(0, 4));
            result = DecodeUtf16WithFastPath(data, 4, len * 2, len);
        }

        return result;
    }

    // ===== Optimized Decoder =====
    /// <summary>
    /// UTF-16 decoder with ASCII-only fast path.
    /// Detects if all characters are ASCII (high byte = 0x00) and uses
    /// optimized path to avoid full UTF-16 decoder overhead.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static string DecodeUtf16WithFastPath(byte[] buffer, int offset, int byteCount, int charCount)
    {
        ReadOnlySpan<byte> span = new ReadOnlySpan<byte>(buffer, offset, byteCount);
        // View bytes as UTF-16 LE code units (ushort)
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);
        // Fast ASCII detection: check if high byte of each code unit is zero
        for (int i = 0; i < units.Length; i++)
        {
            if ((units[i] & 0xFF00) != 0)
            {
                // Non-ASCII detected, fallback to standard decoder
                return Encoding.Unicode.GetString(buffer, offset, byteCount);
            }
        }

        // All characters are ASCII-only. Cast UTF-16 bytes directly to char span.
        // This avoids the overhead of the full UTF-16 decoder.
        ReadOnlySpan<char> chars = MemoryMarshal.Cast<byte, char>(span);
        return new string (chars.Slice(0, charCount));
    }
}