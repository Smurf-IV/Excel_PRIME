using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Numerics;
using System.Runtime.CompilerServices;

using BenchmarkDotNet.Attributes;

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
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = integerSamples[i].Span;
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
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = integerSamples[i].Span;
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
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = integerSamples[i].Span;
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
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = negativeSamples[i].Span;
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
    public object ParseDecimalsCurrent()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = decimalSamples[i].Span;
            bool containsDecimal = span.Contains('.');
            if (containsDecimal)
            {
                if (decimal.TryParse(span, NumberStyles.Currency, invariant, out decimal resultM))
                {
                    result = resultM;
                }
                else if (double.TryParse(span, NumberStyles.Float, invariant, out double resultD))
                {
                    result = resultD;
                }
            }
        }
        return result!;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseDecimalsDoubleFirst()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = decimalSamples[i].Span;
            bool containsDecimal = span.Contains('.');
            if (containsDecimal)
            {
                if (double.TryParse(span, NumberStyles.Float, invariant, out double resultD))
                {
                    result = resultD;
                }
                else if (decimal.TryParse(span, NumberStyles.Currency, invariant, out decimal resultM))
                {
                    result = resultM;
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
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = largeIntegerSamples[i].Span;
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
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = integerSamples[i].Span;
            result = span.Contains('.');
        }
        return result;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseBigIntegersCurrent()
    {
        object? result = null;
        CultureInfo invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            ReadOnlySpan<char> span = bigIntegerSamples[i].Span;
            bool containsDecimal = span.Contains('.');
            if (!containsDecimal && span.Length > 18)
            {
                if (BigInteger.TryParse(span, NumberStyles.Integer, invariant, out BigInteger resultBI))
                {
                    result = resultBI;
                }
            }
        }
        return result!;
    }
}