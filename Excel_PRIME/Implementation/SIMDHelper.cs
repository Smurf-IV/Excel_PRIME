using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace ExcelPRIME.Implementation;

/// <summary>
/// SIMD-accelerated numeric operations for bulk cell processing.
/// Provides vectorized implementations of common filtering and parsing operations
/// with automatic fallback to scalar paths on unsupported architectures.
/// </summary>
internal static class SIMDHelper
{
    /// <summary>
    /// Checks if the CPU supports AVX2 instructions for vectorized double operations.
    /// Used to enable/disable SIMD fast paths at runtime.
    /// </summary>
    private static readonly bool s_supportsAvx2 = Avx2.IsSupported;

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

        if (!s_supportsAvx2)
        {
            return AnyGreaterThan_Scalar(values, threshold);
        }

        int i = 0;
        int vectorCount = Vector256<double>.Count;
        int simdEnd = values.Length - (values.Length % vectorCount);

        Vector256<double> thresholdVec = Vector256.Create(threshold);

        for (; i < simdEnd; i += vectorCount)
        {
            Vector256<double> current = Vector256.Create(
                values[i],
                values[i + 1],
                values[i + 2],
                values[i + 3]
            );

            // Check if any element is greater than threshold
            Vector256<double> gtThreshold = Avx.CompareGreaterThan(current, thresholdVec);

            // If any comparison result is true (all 1 bits), we found a match
            if (Avx.MoveMask(gtThreshold) != 0)
            {
                return true;
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

        if (!s_supportsAvx2)
        {
            return CountGreaterThan_Scalar(values, threshold);
        }

        int count = 0;
        int i = 0;
        int vectorCount = Vector256<double>.Count;
        int simdEnd = values.Length - (values.Length % vectorCount);

        Vector256<double> thresholdVec = Vector256.Create(threshold);

        for (; i < simdEnd; i += vectorCount)
        {
            Vector256<double> current = Vector256.Create(
                values[i],
                values[i + 1],
                values[i + 2],
                values[i + 3]
            );

            Vector256<double> gtThreshold = Avx.CompareGreaterThan(current, thresholdVec);
            int mask = Avx.MoveMask(gtThreshold);

            // Count number of set bits in mask (each bit represents one element > threshold)
            count += PopCount(mask);
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

        if (!s_supportsAvx2)
        {
            return Sum_Scalar(values);
        }

        // Initialize accumulator vector with zeros
        Vector256<double> accumulator = Vector256<double>.Zero;
        int i = 0;
        int vectorCount = Vector256<double>.Count;
        int simdEnd = values.Length - (values.Length % vectorCount);

        // Process 4 doubles at a time
        for (; i < simdEnd; i += vectorCount)
        {
            Vector256<double> current = Vector256.Create(
                values[i],
                values[i + 1],
                values[i + 2],
                values[i + 3]
            );
            accumulator = Avx.Add(accumulator, current);
        }

        // Horizontal sum: add all elements in the accumulator
        // AVX2 doesn't have a direct horizontal add for doubles, so we extract and sum
        double sum = accumulator[0] + accumulator[1] + accumulator[2] + accumulator[3];

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

        if (!s_supportsAvx2)
        {
            return Min_Scalar(values);
        }

        Vector256<double> minVec = Vector256.Create(double.MaxValue);
        int i = 0;
        int vectorCount = Vector256<double>.Count;
        int simdEnd = values.Length - (values.Length % vectorCount);

        for (; i < simdEnd; i += vectorCount)
        {
            Vector256<double> current = Vector256.Create(
                values[i],
                values[i + 1],
                values[i + 2],
                values[i + 3]
            );
            minVec = Avx.Min(minVec, current);
        }

        double min = double.MaxValue;
        for (int j = 0; j < vectorCount; j++)
        {
            min = Math.Min(min, minVec[j]);
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

        if (!s_supportsAvx2)
        {
            return Max_Scalar(values);
        }

        Vector256<double> maxVec = Vector256.Create(double.MinValue);
        int i = 0;
        int vectorCount = Vector256<double>.Count;
        int simdEnd = values.Length - (values.Length % vectorCount);

        for (; i < simdEnd; i += vectorCount)
        {
            Vector256<double> current = Vector256.Create(
                values[i],
                values[i + 1],
                values[i + 2],
                values[i + 3]
            );
            maxVec = Avx.Max(maxVec, current);
        }

        double max = double.MinValue;
        for (int j = 0; j < vectorCount; j++)
        {
            max = Math.Max(max, maxVec[j]);
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
}
