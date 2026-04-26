using System;
using System.Runtime.CompilerServices;

namespace ExcelPRIME.FromExternal;

internal static class Extensions
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IntParse(this string value) => value.AsSpan().IntParse();

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
}
