using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;

namespace ExcelPRIME.Bench;
[ExcludeFromCodeCoverage]
[SimpleJob(warmupCount: 3, iterationCount: 5)]
public class NumericParsingOptimizationBenchmark
{
    private const int NumSamples = 10000;
    private ReadOnlyMemory<char>[] singleDigitSamples = null !;
    private ReadOnlyMemory<char>[] doubleDigitSamples = null !;
    private ReadOnlyMemory<char>[] integerSamples = null !;
    private ReadOnlyMemory<char>[] negativeSamples = null !;
    private ReadOnlyMemory<char>[] decimalSamples = null !;
    private ReadOnlyMemory<char>[] largeIntegerSamples = null !;
    private ReadOnlyMemory<char>[] sciNotationSamples = null !;
    [GlobalSetup]
    public void Setup()
    {
        singleDigitSamples = new ReadOnlyMemory<char>[NumSamples];
        doubleDigitSamples = new ReadOnlyMemory<char>[NumSamples];
        integerSamples = new ReadOnlyMemory<char>[NumSamples];
        negativeSamples = new ReadOnlyMemory<char>[NumSamples];
        decimalSamples = new ReadOnlyMemory<char>[NumSamples];
        largeIntegerSamples = new ReadOnlyMemory<char>[NumSamples];
        sciNotationSamples = new ReadOnlyMemory<char>[NumSamples];
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            // Hot path: single/double digit positive integers (most common in Excel)
            singleDigitSamples[i] = (i % 10).ToString(invariant).AsMemory();
            doubleDigitSamples[i] = (i % 100).ToString(invariant).AsMemory();
            integerSamples[i] = (i * 10).ToString(invariant).AsMemory();
            negativeSamples[i] = (-(i + 1)).ToString(invariant).AsMemory();
            decimalSamples[i] = (i + 0.5f).ToString("F2", invariant).AsMemory();
            largeIntegerSamples[i] = ((long)i * 1000000).ToString(invariant).AsMemory();
            sciNotationSamples[i] = (i * 1.23456789e10).ToString("E6", invariant).AsMemory();
        }
    }

    /// <summary>
    /// Fast int parser for positive ASCII-only digit spans.
    /// Assumes input contains only digits (no sign, no sentinel).
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static int CustomIntParse(ReadOnlySpan<char> value)
    {
        int result = 0;
        int i = 0;
        int valueLength = value.Length - (value.Length & 3);
        // Unrolled 4-digit chunks
        for (; i < valueLength; i += 4)
        {
            result = (10 * result) + (value[i] - '0');
            result = (10 * result) + (value[i + 1] - '0');
            result = (10 * result) + (value[i + 2] - '0');
            result = (10 * result) + (value[i + 3] - '0');
        }

        // Handle remaining digits
        for (; i < value.Length; i++)
        {
            result = (10 * result) + (value[i] - '0');
        }

        return result;
    }

    /// <summary>
    /// Baseline: Original implementation with Contains('.') check
    /// /// </summary>
    //[Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseIntegersBaselineContains()
    {
        object? result = null;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = integerSamples[i].Span;
            bool containsDecimal = span.Contains('.');
            if (!containsDecimal && span.Length < 12)
            {
                if (span[0] != '-')
                {
                    result = CustomIntParse(span);
                }
                else if (int.TryParse(span, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resultI))
                {
                    result = resultI;
                }
            }
        }

        return result!;
    }

    /// <summary>
    /// Optimized: Skip Contains check, use IndexOf for decimal detection only when needed
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseIntegersOptimized()
    {
        object? result = null;
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = integerSamples[i].Span;
            // Direct path for positive integers (no decimal check)
            if (span.Length < 12 && span[0] != '-')
            {
                result = CustomIntParse(span);
            }
            else if (span.IndexOf('.') < 0 && span.Length < 20)
            {
                if (int.TryParse(span, NumberStyles.Integer, invariant, out int resultI))
                {
                    result = resultI;
                }
                else if (long.TryParse(span, NumberStyles.Integer, invariant, out long resultL))
                {
                    result = resultL;
                }
            }
        }

        return result!;
    }

    /// <summary>
    /// Single digit integer - tests the single digit fast path
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseSingleDigitIntegers()
    {
        object? result = null;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = singleDigitSamples[i].Span;
            // Single digit fast path
            if (span.Length == 1 && span[0] >= '0' && span[0] <= '9')
            {
                result = span[0] - '0';
            }
        }

        return result!;
    }

    /// <summary>
    /// Double digit integer - tests double digit handling
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseDoubleDigitIntegers()
    {
        object? result = null;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = doubleDigitSamples[i].Span;
            if (span.Length < 12 && span[0] != '-')
            {
                result = CustomIntParse(span);
            }
        }

        return result!;
    }

    /// <summary>
    /// Negative integers - tests negative number path
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseNegativeIntegers()
    {
        object? result = null;
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = negativeSamples[i].Span;
            if (span.Length < 12 && span[0] == '-')
            {
                if (int.TryParse(span, NumberStyles.Integer, invariant, out int resultI))
                {
                    result = resultI;
                }
            }
        }

        return result!;
    }

    /// <summary>
    /// Decimal with double.TryParse first (optimized order)
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseDecimalsDoubleFirst()
    {
        object? result = null;
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = decimalSamples[i].Span;
            if (span.IndexOf('.') >= 0)
            {
                if (double.TryParse(span, NumberStyles.Float, invariant, out double resultD))
                {
                    result = resultD;
                }
                else if (decimal.TryParse(span, NumberStyles.Number, invariant, out decimal resultM))
                {
                    result = resultM;
                }
            }
        }

        return result!;
    }

    /// <summary>
    /// Decimal with decimal.TryParse first (baseline order)
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseDecimalsDecimalFirst()
    {
        object? result = null;
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = decimalSamples[i].Span;
            if (span.IndexOf('.') >= 0)
            {
                if (decimal.TryParse(span, NumberStyles.Number, invariant, out decimal resultM))
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

    /// <summary>
    /// Large integers with long.TryParse
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseLargeIntegers()
    {
        object? result = null;
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = largeIntegerSamples[i].Span;
            if (span.Length < 20)
            {
                if (long.TryParse(span, NumberStyles.Integer, invariant, out long resultL))
                {
                    result = resultL;
                }
            }
        }

        return result!;
    }

    /// <summary>
    /// Scientific notation with double
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public object ParseScientificNotation()
    {
        object? result = null;
        var invariant = CultureInfo.InvariantCulture;
        for (int i = 0; i < NumSamples; i++)
        {
            var span = sciNotationSamples[i].Span;
            if (double.TryParse(span, NumberStyles.Float, invariant, out double resultD))
            {
                result = resultD;
            }
        }

        return result!;
    }
}