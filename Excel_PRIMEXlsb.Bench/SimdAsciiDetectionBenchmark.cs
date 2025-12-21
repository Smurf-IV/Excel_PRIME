using System;
using System.Runtime.InteropServices;
using System.Numerics;
using System.Diagnostics.CodeAnalysis;
using System.Text;

using BenchmarkDotNet.Attributes;

namespace ExcelPRIMEXlsb.Bench;

/*
| Method                         | Job        | IterationCount | LaunchCount | Ratio        | Allocated | Alloc Ratio |
   |------------------------------- |----------- |--------------- |------------ |-------------:|----------:|------------:|
   | ScalarBatch4AsciiDetection     | Job-NTRUNJ | 5              | Default     |     baseline |         - |          NA |
   | SimdVectorUshortAsciiDetection | Job-NTRUNJ | 5              | Default     | 1.74x faster |         - |          NA |
   | SimdVectorByteAsciiDetection   | Job-NTRUNJ | 5              | Default     | 1.10x slower |         - |          NA |
   | ScalarBatch4LatinDetection     | Job-NTRUNJ | 5              | Default     | 1.00x faster |         - |          NA |
   | SimdVectorUshortLatinDetection | Job-NTRUNJ | 5              | Default     | 1.78x faster |         - |          NA |
   | SimdVectorByteLatinDetection   | Job-NTRUNJ | 5              | Default     | 1.10x slower |         - |          NA |
   | ScalarBatch4ShortDetection     | Job-NTRUNJ | 5              | Default     | 2.59x faster |         - |          NA |
   | SimdVectorUshortShortDetection | Job-NTRUNJ | 5              | Default     | 2.01x faster |         - |          NA |
   | SimdVectorByteShortDetection   | Job-NTRUNJ | 5              | Default     | 3.60x faster |         - |          NA |
   |                                |            |                |             |              |           |             |
   | ScalarBatch4AsciiDetection     | ShortRun   | 3              | 1           |     baseline |         - |          NA |
   | SimdVectorUshortAsciiDetection | ShortRun   | 3              | 1           | 1.78x faster |         - |          NA |
   | SimdVectorByteAsciiDetection   | ShortRun   | 3              | 1           | 1.11x slower |         - |          NA |
   | ScalarBatch4LatinDetection     | ShortRun   | 3              | 1           | 1.00x slower |         - |          NA |
   | SimdVectorUshortLatinDetection | ShortRun   | 3              | 1           | 1.73x faster |         - |          NA |
   | SimdVectorByteLatinDetection   | ShortRun   | 3              | 1           | 1.12x slower |         - |          NA |
   | ScalarBatch4ShortDetection     | ShortRun   | 3              | 1           | 2.56x faster |         - |          NA |
   | SimdVectorUshortShortDetection | ShortRun   | 3              | 1           | 2.00x faster |         - |          NA |
   | SimdVectorByteShortDetection   | ShortRun   | 3              | 1           | 3.60x faster |         - |          NA |
 */
/// <summary>
/// Benchmark for SIMD vs scalar ASCII detection in UTF-16 strings.
/// Compares performance of different vectorization strategies.
/// </summary>
[ExcludeFromCodeCoverage]
[SimpleJob(warmupCount: 3, iterationCount: 5)]
public class SimdAsciiDetectionBenchmark
{
    private const int StringCount = 1000;
    private byte[][] asciiStrings = null!;
    private byte[][] latinStrings = null!;
    private byte[][] shortAsciiStrings = null!;

    [GlobalSetup]
    public void Setup()
    {
        // ASCII strings (5-50 chars, typical for column names/sheet names)
        asciiStrings = new byte[StringCount][];
        for (int i = 0; i < StringCount; i++)
        {
            string text = GenerateAsciiString(10 + (i % 40));
            asciiStrings[i] = EncodeUtf16String(text);
        }

        // Latin-1 Extended (mix of ASCII and non-ASCII)
        latinStrings = new byte[StringCount][];
        for (int i = 0; i < StringCount; i++)
        {
            string text = (i % 10 == 0) 
                ? GenerateLatin1String(10 + (i % 40))
                : GenerateAsciiString(10 + (i % 40));
            latinStrings[i] = EncodeUtf16String(text);
        }

        // Short ASCII strings (1-10 chars)
        shortAsciiStrings = new byte[StringCount][];
        for (int i = 0; i < StringCount; i++)
        {
            string text = GenerateAsciiString(1 + (i % 10));
            shortAsciiStrings[i] = EncodeUtf16String(text);
        }
    }

    private static string GenerateAsciiString(int length)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < length; i++)
        {
            sb.Append((char)('A' + (i % 26)));
        }
        return sb.ToString();
    }

    private static string GenerateLatin1String(int length)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < length; i++)
        {
            // Mix ASCII and Latin-1 Extended
            if (i % 10 == 0)
            {
                sb.Append((char)(0xC0 + (i % 32))); // À-ß range
            }
            else
            {
                sb.Append((char)('A' + (i % 26)));
            }
        }
        return sb.ToString();
    }

    private static byte[] EncodeUtf16String(string text)
    {
        byte[] encoded = Encoding.Unicode.GetBytes(text);
        byte[] result = new byte[encoded.Length + 4];
        BitConverter.GetBytes(text.Length).CopyTo(result, 0);
        encoded.CopyTo(result, 4);
        return result;
    }

    // ===== Scalar Baseline (4-unit batch) =====

    //[Benchmark(Baseline = true)]
    [BenchmarkCategory("ASCII")]
    public int ScalarBatch4AsciiDetection()
    {
        int result = 0;
        foreach (byte[] data in asciiStrings)
        {
            if (IsAsciiOnlyScalarBatch4(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    //[Benchmark]
    [BenchmarkCategory("Latin")]
    public int ScalarBatch4LatinDetection()
    {
        int result = 0;
        foreach (byte[] data in latinStrings)
        {
            if (IsAsciiOnlyScalarBatch4(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    //[Benchmark]
    [BenchmarkCategory("Short")]
    public int ScalarBatch4ShortDetection()
    {
        int result = 0;
        foreach (byte[] data in shortAsciiStrings)
        {
            if (IsAsciiOnlyScalarBatch4(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    // ===== SIMD Vectorized (Vector<ushort>) =====

    //[Benchmark]
    [BenchmarkCategory("ASCII")]
    public int SimdVectorUshortAsciiDetection()
    {
        int result = 0;
        foreach (byte[] data in asciiStrings)
        {
            if (IsAsciiOnlySimdVectorUshort(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    //[Benchmark]
    [BenchmarkCategory("Latin")]
    public int SimdVectorUshortLatinDetection()
    {
        int result = 0;
        foreach (byte[] data in latinStrings)
        {
            if (IsAsciiOnlySimdVectorUshort(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    //[Benchmark]
    [BenchmarkCategory("Short")]
    public int SimdVectorUshortShortDetection()
    {
        int result = 0;
        foreach (byte[] data in shortAsciiStrings)
        {
            if (IsAsciiOnlySimdVectorUshort(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    // ===== SIMD Byte-based (Vector<byte>) =====

    //[Benchmark]
    [BenchmarkCategory("ASCII")]
    public int SimdVectorByteAsciiDetection()
    {
        int result = 0;
        foreach (byte[] data in asciiStrings)
        {
            if (IsAsciiOnlySimdVectorByte(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    //[Benchmark]
    [BenchmarkCategory("Latin")]
    public int SimdVectorByteLatinDetection()
    {
        int result = 0;
        foreach (byte[] data in latinStrings)
        {
            if (IsAsciiOnlySimdVectorByte(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    //[Benchmark]
    [BenchmarkCategory("Short")]
    public int SimdVectorByteShortDetection()
    {
        int result = 0;
        foreach (byte[] data in shortAsciiStrings)
        {
            if (IsAsciiOnlySimdVectorByte(data, 4))
            {
                result++;
            }
        }
        return result;
    }

    // ===== Implementation Methods =====

    private static bool IsAsciiOnlyScalarBatch4(byte[] buffer, int offset)
    {
        int len = BitConverter.ToInt32(buffer, 0);
        if (len <= 0)
        {
            return true;
        }

        ReadOnlySpan<byte> span = new ReadOnlySpan<byte>(buffer, offset, len * 2);
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);

        int i = 0;
        int unitsLength = units.Length;
        const int BATCH_SIZE = 4;
        int batchEnd = unitsLength - (unitsLength % BATCH_SIZE);

        for (; i < batchEnd; i += BATCH_SIZE)
        {
            if (((units[i] | units[i + 1] | units[i + 2] | units[i + 3]) & 0xFF00) != 0)
            {
                return false;
            }
        }

        for (; i < unitsLength; i++)
        {
            if ((units[i] & 0xFF00) != 0)
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsAsciiOnlySimdVectorUshort(byte[] buffer, int offset)
    {
        int len = BitConverter.ToInt32(buffer, 0);
        if (len <= 0)
        {
            return true;
        }

        ReadOnlySpan<byte> span = new ReadOnlySpan<byte>(buffer, offset, len * 2);
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);

        int i = 0;
        int vectorSize = Vector<ushort>.Count;
        int unitsLength = units.Length;
        int simdEnd = unitsLength - (unitsLength % vectorSize);

        // SIMD fast path: check multiple code units in parallel
        Vector<ushort> mask = new Vector<ushort>(0xFF00);
        for (; i < simdEnd; i += vectorSize)
        {
            Vector<ushort> vector = new Vector<ushort>(units.Slice(i, vectorSize));
            Vector<ushort> result = vector & mask;
            
            if (!Vector.EqualsAll(result, Vector<ushort>.Zero))
            {
                return false;
            }
        }

        // Handle remaining units
        for (; i < unitsLength; i++)
        {
            if ((units[i] & 0xFF00) != 0)
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsAsciiOnlySimdVectorByte(byte[] buffer, int offset)
    {
        int len = BitConverter.ToInt32(buffer, 0);
        if (len <= 0)
        {
            return true;
        }

        ReadOnlySpan<byte> span = new ReadOnlySpan<byte>(buffer, offset, len * 2);

        // Fall back to scalar - Vector<byte> approach is complex for this use case
        ReadOnlySpan<ushort> units = MemoryMarshal.Cast<byte, ushort>(span);
        for (int i = 0; i < units.Length; i++)
        {
            if ((units[i] & 0xFF00) != 0)
            {
                return false;
            }
        }

        return true;
    }
}
