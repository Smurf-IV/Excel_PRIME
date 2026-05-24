using System;
using System.Collections.Generic;
using System.Globalization;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelPRIME.FromExternal;

internal static class Extensions
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string ToInvariantString<T>(this T value) where T : INumber<T>, IFormattable
        => value.ToString("R", CultureInfo.InvariantCulture);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IntParse(this string value) => value.AsSpan().IntParse();

    /// <summary>
    /// Optimized boolean parsing from span without allocation.
    /// Returns true if span represents '1' or 'true', false for '0' or 'false'.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool BoolParse(this ReadOnlySpan<char> value) =>
        value.Length switch
        {
            0 => false,
            1 => value[0] == '1',
            4 => value is "true" || value is "True" || value is "TRUE",
            5 => value is "false" || value is "False" || value is "FALSE",
            _ => false
        };

    /// <summary>
    /// Optimized double parsing from span.
    /// Wrapper around double.TryParse with ReadOnlySpan for performance.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool TryDoubleParse(this ReadOnlySpan<char> value, out double result)
        => double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out result);

    /// <summary>
    /// Optimized integer parsing with loop unrolling for .NET 8+ platforms.
    /// Uses process 4 characters at a time with manual bounds checking to reduce allocations
    /// and improve CPU cache efficiency. Compatible with custom XML parsing scenarios.
    /// 
    /// Performance improvement in .NET 8: Custom parsing remains faster than int.Parse(ReadOnlySpan)
    /// due to bounds-check optimization and inline-friendly control flow.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static int IntParse(this ReadOnlySpan<char> value)
    {
        int result = 0;
        int i = 0;
        // outside the for loop to allow bounds-check once
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
        // Do the rest
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

    // ParseDecimal taken from https://stackoverflow.com/a/37754822
    // And then modified to be faster and allocate less
    private static readonly int[] s_powOf10 =
    [
        1,
        10,
        100,
        1000,
        10000,
        100000,
        1000000,
        10000000,
        100000000,
        1000000000
    ];
    /// <summary>
    /// Optimized decimal parsing from span.
    /// Wrapper around decimal.TryParse with ReadOnlySpan for performance.
    /// </summary>
    public static bool TryDecimalParse(this ReadOnlySpan<char> input, out decimal result)
    {
        int len = input.Length;
        result = 0;
        if (len == 0)
        {
            return false;
        }

        bool negative = false;
        long n = 0;
        int start = 0;
        if (input[0] == '-')
        {
            negative = true;
            start = 1;
        }
        int decPos = len;
        if (len <= 19)
        {
            for (int k = start; k < len; k++)
            {
                ref readonly char c = ref input[k];
                if (c == '.')
                {
                    decPos = k + 1;
                }
                else if(char.IsBetween(c, '0', '9'))
                {
                    n = (n * 10) + (int)(c - '0');
                }
                else
                {
                    return false;
                }
            }
            result = new decimal((int)n, (int)(n >> 32), 0, negative, (byte)(len - decPos));
            return true;
        }

        if (len > 28)
        {
            len = 28;
        }
        for (int k = start; k < 19; k++)
        {
            ref readonly char c = ref input[k];
            if (c == '.')
            {
                decPos = k + 1;
            }
            else if (char.IsBetween(c, '0', '9'))
            {
                n = (n * 10) + (int)(c - '0');
            }
            else
            {
                return false;
            }
        }
        int n2 = 0;
        bool secondHalfDec = false;
        for (int k = 19; k < len; k++)
        {
            ref readonly char c = ref input[k];
            if (c == '.')
            {
                decPos = k + 1;
                secondHalfDec = true;
            }
            else if (char.IsBetween(c, '0', '9'))
            {
                n2 = (n2 * 10) + (int)(c - '0');
            }
            else
            {
                return false;
            }
        }
        byte decimalPosition = (byte)(len - decPos);
        result = new decimal((int)n, (int)(n >> 32), 0, negative, decimalPosition) * s_powOf10[len - (!secondHalfDec ? 19 : 20)] + new decimal(n2, 0, 0, negative, decimalPosition);
        return true;
    }

    /// <summary>
    /// Zero-allocation GetEnumerator for Span without struct wrapping.
    /// Returns direct span enumerator which iterates without allocation.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static Span<T>.Enumerator GetEnumerator<T>(this Span<T> span) => span.GetEnumerator();

    /// <summary>
    /// Zero-allocation AsSpan extension for List{T}.
    /// Uses CollectionsMarshal to access underlying array without copying.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static Span<T> AsSpan<T>(this List<T> list) => CollectionsMarshal.AsSpan(list);

    /// <summary>
    /// Zero-allocation Any predicate check on Span{T}.
    /// Returns true if any element matches the predicate without allocation.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool Any<T>(this Span<T> span, Predicate<T> predicate)
    {
        foreach (T item in span)
        {
            if (predicate(item))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Zero-allocation All predicate check on Span{T}.
    /// Returns true if all elements match the predicate without allocation.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool All<T>(this Span<T> span, Predicate<T> predicate)
    {
        foreach (T item in span)
        {
            if (!predicate(item))
            {
                return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Zero-allocation Count predicate check on Span{T}.
    /// Counts elements matching the predicate without allocation.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int Count<T>(this Span<T> span, Predicate<T> predicate)
    {
        int count = 0;
        foreach (T item in span)
        {
            if (predicate(item))
            {
                count++;
            }
        }
        return count;
    }

    /// <summary>
    /// Zero-allocation filter operation on Span{T} into a pre-allocated output span.
    /// Copies all matching elements to output span and returns the count of items written.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int Filter<T>(this Span<T> input, Span<T> output, Predicate<T> predicate)
    {
        int outputIndex = 0;
        foreach (T item in input)
        {
            if (predicate(item))
            {
                if (outputIndex < output.Length)
                {
                    output[outputIndex++] = item;
                }
                else
                {
                    break;
                }
            }
        }
        return outputIndex;
    }

    /// <summary>
    /// Zero-allocation filter operation on ReadOnlySpan{T} into a pre-allocated output span.
    /// Copies all matching elements to output span and returns the count of items written.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int Filter<T>(this ReadOnlySpan<T> input, Span<T> output, Predicate<T> predicate)
    {
        int outputIndex = 0;
        foreach (T item in input)
        {
            if (predicate(item))
            {
                if (outputIndex < output.Length)
                {
                    output[outputIndex++] = item;
                }
                else
                {
                    break;
                }
            }
        }
        return outputIndex;
    }

    /// <summary>
    /// Zero-allocation first match operation on Span{T}.
    /// Returns the first element matching the predicate, or default(T) if none found.
    /// Sets found parameter to indicate if a match was found.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static T FirstOrDefault<T>(this Span<T> span, Predicate<T> predicate, out bool found)
    {
        foreach (T item in span)
        {
            if (predicate(item))
            {
                found = true;
                return item;
            }
        }
        found = false;
        return default!;
    }

    /// <summary>
    /// Zero-allocation first match operation on ReadOnlySpan{T}.
    /// Returns the first element matching the predicate, or default(T) if none found.
    /// Sets found parameter to indicate if a match was found.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static T FirstOrDefault<T>(this ReadOnlySpan<T> span, Predicate<T> predicate, out bool found)
    {
        foreach (T item in span)
        {
            if (predicate(item))
            {
                found = true;
                return item;
            }
        }
        found = false;
        return default!;
    }
}
