using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Numerics;
using System.Runtime.CompilerServices;

using BenchmarkDotNet.Attributes;
using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Bench;

[ExcludeFromCodeCoverage]
[SimpleJob(warmupCount: 2, iterationCount: 5)]
[MemoryDiagnoser]
public class NumericParsingBenchmark
{
    private const int NumSamples = 10000;
    private ReadOnlyMemory<char>[] integerSamples = null!;
    private ReadOnlyMemory<char>[] negativeSamples = null!;
    private ReadOnlyMemory<char>[] decimalSamples = null!;
    private ReadOnlyMemory<char>[] largeIntegerSamples = null!;
    private ReadOnlyMemory<char>[] bigIntegerSamples = null!;

    [GlobalSetup]
    public void Setup()
    {
        integerSamples = new ReadOnlyMemory<char>[NumSamples];
        negativeSamples = new ReadOnlyMemory<char>[NumSamples];
        decimalSamples = new ReadOnlyMemory<char>[NumSamples];
        largeIntegerSamples = new ReadOnlyMemory<char>[NumSamples];
        bigIntegerSamples = new ReadOnlyMemory<char>[NumSamples];

        for (int i = 0; i < NumSamples; i++)
        {
            integerSamples[i] = i.ToString().AsMemory();
            negativeSamples[i] = (-i).ToString().AsMemory();
            decimalSamples[i] = (i + 0.5f).ToString("F2").AsMemory();
            largeIntegerSamples[i] = ((long)i * 1000000).ToString().AsMemory();
            bigIntegerSamples[i] = ((BigInteger)i * BigInteger.Parse("999999999999999")).ToString().AsMemory();
        }
    }

    /// <summary>
    /// Custom fast int parser - simulates the internal IntParse extension
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static int CustomIntParse(ReadOnlySpan<char> value)
    {
        int result = 0;
        int i = 0;
        int valueLength = value.Length - (value.Length & 3);

        for (; i < valueLength; i += 4)
        {
            ref readonly char local = ref value[i];
            if (local != '\0')
            {
                result = (10 * result) + (local - 48);
            }
            else
            {
                i = value.Length;
                break;
            }
            local = ref value[i + 1];
            if (local != '\0')
            {
                result = (10 * result) + (local - 48);
            }
            else
            {
                i = value.Length;
                break;
            }
            local = ref value[i + 2];
            if (local != '\0')
            {
                result = (10 * result) + (local - 48);
            }
            else
            {
                i = value.Length;
                break;
            }
            local = ref value[i + 3];
            if (local != '\0')
            {
                result = (10 * result) + (local - 48);
            }
            else
            {
                i = value.Length;
                break;
            }
        }

        for (; i < value.Length; i++)
        {
            ref readonly char local = ref value[i];
            if (local != '\0')
            {
                result = (10 * result) + (local - 48);
            }
            else
            {
                break;
            }
        }

        return result;
    }

    //[Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseIntegersCurrent()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        foreach (ReadOnlyMemory<char> sample in integerSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            bool containsDecimal = span.Contains('.');
            if (!containsDecimal && span.Length < 12)
            {
                if (span[0] != '-')
                {
                    result = CustomIntParse(span);
                }
                else if (int.TryParse(span, NumberStyles.Integer, invariant, out int resultI))
                {
                    result = resultI;
                }
            }
        }
        return result!;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseIntegersOptimized()
    {
        object? result = null;
        foreach (ReadOnlyMemory<char> sample in integerSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            if (span.Length < 12 && span[0] != '-')
            {
                result = CustomIntParse(span);
            }
        }
        return result!;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseIntegersTryParse()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        foreach (ReadOnlyMemory<char> sample in integerSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            if (int.TryParse(span, NumberStyles.Integer, invariant, out int resultI))
            {
                result = resultI;
            }
        }
        return result!;
    }

   //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseNegativeIntegersCurrent()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        foreach (ReadOnlyMemory<char> sample in negativeSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            bool containsDecimal = span.Contains('.');
            if (!containsDecimal && span.Length < 12)
            {
                if (span[0] != '-')
                {
                    result = CustomIntParse(span);
                }
                else if (int.TryParse(span, NumberStyles.Integer, invariant, out int resultI))
                {
                    result = resultI;
                }
            }
        }
        return result!;
    }



    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseLargeIntegersCurrent()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        foreach (ReadOnlyMemory<char> sample in largeIntegerSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            bool containsDecimal = span.Contains('.');
            if (!containsDecimal && span.Length < 20)
            {
                if (long.TryParse(span, NumberStyles.Integer, invariant, out long resultL))
                {
                    result = resultL;
                }
            }
        }
        return result!;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public bool MeasureContainsCheckCost()
    {
        bool result = false;
        foreach (ReadOnlyMemory<char> sample in integerSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            result = span.Contains('.');
        }
        return result;
    }


    /// <summary>
    /// Benchmark for optimized DecimalParse extension method.
    /// Measures performance of the custom zero-allocation decimal parsing using DecimalParse from Extensions.
    /// Uses an optimized algorithm with minimal allocations for parsing decimal values.
    /// </summary>
    /*
    | Method          | Job        | IterationCount | LaunchCount | WarmupCount | Ratio        | Allocated | Alloc Ratio |
    |---------------- |----------- |--------------- |------------ |------------ |-------------:|----------:|------------:|
    | DecimalParse    | Job-AMZPBM | 5              | Default     | 2           | 2.71x faster |         - |          NA |
    | TryDecimalParse | Job-AMZPBM | 5              | Default     | 2           |     baseline |         - |          NA |
    */
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public decimal DecimalParse()
    {
        decimal result = 0;
        foreach (ReadOnlyMemory<char> sample in decimalSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            try
            {
                span.TryDecimalParse(out result);
            }
            catch
            {
                // Skip parsing errors for invalid formats
            }
        }
        return result;
    }

    //[Benchmark(Baseline=true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public decimal TryDecimalParse()
    {
        decimal result = 0;
        foreach (ReadOnlyMemory<char> sample in decimalSamples)
        {
            ReadOnlySpan<char> span = sample.Span;
            try
            {
                decimal.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
            }
            catch
            {
                // Skip parsing errors for invalid formats
            }
        }
        return result;
    }

}