using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;

namespace ExcelPRIME.Implementation;

/// <summary>
/// SIMD-accelerated numeric operations for bulk cell processing.
/// Provides vectorized implementations of common filtering and parsing operations
/// with automatic fallback to scalar paths on unsupported architectures.
/// </summary>
internal static class SIMDHelper
{
    /// <summary>
    /// Checks if the CPU supports vectorized operations.
    /// Used to enable/disable SIMD fast paths at runtime.
    /// </summary>
    private static readonly bool s_supportsVector512 = Vector512.IsHardwareAccelerated;
    private static readonly bool s_supportsVector256 = Vector256.IsHardwareAccelerated;

    /// <summary>
    /// SIMD-accelerated check if any value in the input exceeds a threshold.
    /// Processes multiple doubles in parallel using AVX2.
    /// </summary>
    /// <param name="values">Input array of numeric values.</param>
    /// <param name="threshold">The threshold value.</param>
    /// <returns>True if any value is greater than threshold, false otherwise.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static bool AnyGreaterThan(ReadOnlySpan<double> values, double threshold)
    {
        if (values.IsEmpty)
        {
            return false;
        }

        int i = 0;
        Vector256<double> thresholdVec256 = Vector256.Create(threshold);
        ref double valuesRef = ref MemoryMarshal.GetReference(values);

        if (s_supportsVector512 && values.Length >= Vector512<double>.Count)
        {
            int vectorCount = Vector512<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector512<double> thresholdVec = Vector512.Create(threshold);

            for (; i < simdEnd; i += vectorCount)
            {
                Vector512<double> current = Vector512.LoadUnsafe(ref valuesRef, (uint)i);
#if NET9_0_OR_GREATER
                if (Vector512.GreaterThanAny(current, thresholdVec))
#else
                if (Vector512.GreaterThan(current, thresholdVec) != Vector512<double>.Zero)
#endif
                {
                    return true;
                }
            }
        }

        if (s_supportsVector256 && values.Length - i >= Vector256<double>.Count)
        {
            int vectorCount = Vector256<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);

            for (; i < simdEnd; i += vectorCount)
            {
                Vector256<double> current = Vector256.LoadUnsafe(ref valuesRef, (uint)i);
                if (Vector256.GreaterThanAny(current, thresholdVec256))
                {
                    return true;
                }
            }
        }

        // Scalar fallback for remaining elements
        for (; i < values.Length; i++)
        {
            if (values[i] > threshold)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Scalar fallback for AnyGreaterThan.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool AnyGreaterThan_Scalar(ReadOnlySpan<double> values, double threshold)
    {
        foreach (double value in values)
        {
            if (value > threshold)
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// SIMD-accelerated count of values meeting a numeric comparison condition.
    /// Counts elements greater than the threshold using vectorized operations.
    /// </summary>
    /// <param name="values">Input array of numeric values.</param>
    /// <param name="threshold">The threshold value.</param>
    /// <returns>Count of values greater than threshold.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static int CountGreaterThan(ReadOnlySpan<double> values, double threshold)
    {
        if (values.IsEmpty)
        {
            return 0;
        }

        int count = 0;
        int i = 0;
        ref double valuesRef = ref MemoryMarshal.GetReference(values);

        if (s_supportsVector512 && values.Length >= Vector512<double>.Count)
        {
            int vectorCount = Vector512<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector512<double> thresholdVec = Vector512.Create(threshold);

            for (; i < simdEnd; i += vectorCount)
            {
                Vector512<double> current = Vector512.LoadUnsafe(ref valuesRef, (uint)i);
                Vector512<double> gtThreshold = Vector512.GreaterThan(current, thresholdVec);
                ulong mask = gtThreshold.ExtractMostSignificantBits();
                count += System.Numerics.BitOperations.PopCount(mask);
            }
        }

        if (s_supportsVector256 && values.Length - i >= Vector256<double>.Count)
        {
            int vectorCount = Vector256<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector256<double> thresholdVec = Vector256.Create(threshold);

            for (; i < simdEnd; i += vectorCount)
            {
                Vector256<double> current = Vector256.LoadUnsafe(ref valuesRef, (uint)i);
                Vector256<double> gtThreshold = Vector256.GreaterThan(current, thresholdVec);
                uint mask = gtThreshold.ExtractMostSignificantBits();
                count += System.Numerics.BitOperations.PopCount(mask);
            }
        }

        // Scalar fallback for remaining elements
        for (; i < values.Length; i++)
        {
            if (values[i] > threshold)
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>
    /// Scalar fallback for CountGreaterThan.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int CountGreaterThan_Scalar(ReadOnlySpan<double> values, double threshold)
    {
        int count = 0;
        foreach (double value in values)
        {
            if (value > threshold)
            {
                count++;
            }
        }
        return count;
    }

    /// <summary>
    /// Count the number of set bits in an integer (population count).
    /// Used to quickly count comparison results from SIMD operations.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int PopCount(int value) =>
        // Use BitOperations.PopCount on .NET 6+, or manual implementation
        System.Numerics.BitOperations.PopCount((uint)value);

    /// <summary>
    /// SIMD-accelerated sum of all values in the input array.
    /// Processes multiple doubles in parallel using AVX2.
    /// </summary>
    /// <param name="values">Input array of numeric values.</param>
    /// <returns>Sum of all values.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static double Sum(ReadOnlySpan<double> values)
    {
        if (values.IsEmpty)
        {
            return 0.0;
        }

        int i = 0;
        double sum = 0.0;
        ref double valuesRef = ref MemoryMarshal.GetReference(values);

        if (s_supportsVector512 && values.Length >= Vector512<double>.Count)
        {
            int vectorCount = Vector512<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector512<double> accumulator = Vector512<double>.Zero;

            for (; i < simdEnd; i += vectorCount)
            {
                accumulator += Vector512.LoadUnsafe(ref valuesRef, (uint)i);
            }
            sum = HorizontalSum(accumulator);
        }

        if (s_supportsVector256 && values.Length - i >= Vector256<double>.Count)
        {
            int vectorCount = Vector256<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector256<double> accumulator = Vector256<double>.Zero;

            for (; i < simdEnd; i += vectorCount)
            {
                accumulator += Vector256.LoadUnsafe(ref valuesRef, (uint)i);
            }
            sum += HorizontalSum(accumulator);
        }

        // Scalar fallback for remaining elements
        for (; i < values.Length; i++)
        {
            sum += values[i];
        }

        return sum;
    }

    /// <summary>
    /// Scalar fallback for Sum.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double Sum_Scalar(ReadOnlySpan<double> values)
    {
        double sum = 0.0;
        foreach (double value in values)
        {
            sum += value;
        }
        return sum;
    }

    /// <summary>
    /// SIMD-accelerated minimum value finder.
    /// </summary>
    /// <param name="values">Input array of numeric values.</param>
    /// <returns>Minimum value in the array, or double.MaxValue if empty.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static double Min(ReadOnlySpan<double> values)
    {
        if (values.IsEmpty)
        {
            return double.MaxValue;
        }

        int i = 0;
        double min = double.MaxValue;
        ref double valuesRef = ref MemoryMarshal.GetReference(values);

        if (s_supportsVector512 && values.Length >= Vector512<double>.Count)
        {
            int vectorCount = Vector512<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector512<double> minVec = Vector512.Create(double.MaxValue);

            for (; i < simdEnd; i += vectorCount)
            {
                minVec = Vector512.Min(minVec, Vector512.LoadUnsafe(ref valuesRef, (uint)i));
            }
            min = HorizontalMin(minVec);
        }

        if (s_supportsVector256 && values.Length - i >= Vector256<double>.Count)
        {
            int vectorCount = Vector256<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector256<double> minVec = Vector256.Create(double.MaxValue);

            for (; i < simdEnd; i += vectorCount)
            {
                minVec = Vector256.Min(minVec, Vector256.LoadUnsafe(ref valuesRef, (uint)i));
            }
            min = Math.Min(min, HorizontalMin(minVec));
        }

        // Scalar fallback for remaining elements
        for (; i < values.Length; i++)
        {
            min = Math.Min(min, values[i]);
        }

        return min;
    }

    /// <summary>
    /// Scalar fallback for Min.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double Min_Scalar(ReadOnlySpan<double> values)
    {
        double min = double.MaxValue;
        foreach (double value in values)
        {
            min = Math.Min(min, value);
        }
        return min;
    }

    /// <summary>
    /// SIMD-accelerated maximum value finder.
    /// </summary>
    /// <param name="values">Input array of numeric values.</param>
    /// <returns>Maximum value in the array, or double.MinValue if empty.</returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static double Max(ReadOnlySpan<double> values)
    {
        if (values.IsEmpty)
        {
            return double.MinValue;
        }

        int i = 0;
        double max = double.MinValue;
        ref double valuesRef = ref MemoryMarshal.GetReference(values);

        if (s_supportsVector512 && values.Length >= Vector512<double>.Count)
        {
            int vectorCount = Vector512<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector512<double> maxVec = Vector512.Create(double.MinValue);

            for (; i < simdEnd; i += vectorCount)
            {
                maxVec = Vector512.Max(maxVec, Vector512.LoadUnsafe(ref valuesRef, (uint)i));
            }
            max = HorizontalMax(maxVec);
        }

        if (s_supportsVector256 && values.Length - i >= Vector256<double>.Count)
        {
            int vectorCount = Vector256<double>.Count;
            int simdEnd = values.Length - (values.Length % vectorCount);
            Vector256<double> maxVec = Vector256.Create(double.MinValue);

            for (; i < simdEnd; i += vectorCount)
            {
                maxVec = Vector256.Max(maxVec, Vector256.LoadUnsafe(ref valuesRef, (uint)i));
            }
            max = Math.Max(max, HorizontalMax(maxVec));
        }

        // Scalar fallback for remaining elements
        for (; i < values.Length; i++)
        {
            max = Math.Max(max, values[i]);
        }

        return max;
    }

    /// <summary>
    /// Scalar fallback for Max.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double Max_Scalar(ReadOnlySpan<double> values)
    {
        double max = double.MinValue;
        foreach (double value in values)
        {
            max = Math.Max(max, value);
        }
        return max;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double HorizontalSum(Vector512<double> v)
    {
#if NET9_0_OR_GREATER
        return Vector512.Sum(v);
#else
        return HorizontalSum(v.GetLower() + v.GetUpper());
#endif
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double HorizontalSum(Vector256<double> v)
    {
#if NET9_0_OR_GREATER
        return Vector256.Sum(v);
#else
        Vector128<double> sum128 = v.GetLower() + v.GetUpper();
        return sum128.GetElement(0) + sum128.GetElement(1);
#endif
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double HorizontalMin(Vector512<double> v)
    {
        return HorizontalMin(Vector256.Min(v.GetLower(), v.GetUpper()));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double HorizontalMin(Vector256<double> v)
    {
        Vector128<double> min128 = Vector128.Min(v.GetLower(), v.GetUpper());
        return Math.Min(min128.GetElement(0), min128.GetElement(1));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double HorizontalMax(Vector512<double> v)
    {
        return HorizontalMax(Vector256.Max(v.GetLower(), v.GetUpper()));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double HorizontalMax(Vector256<double> v)
    {
        Vector128<double> max128 = Vector128.Max(v.GetLower(), v.GetUpper());
        return Math.Max(max128.GetElement(0), max128.GetElement(1));
    }
}
