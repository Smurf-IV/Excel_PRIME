using System;
using System.Collections.Generic;
using System.Globalization;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;


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
    /// Performance improvement in .NET 8: Custom parsing remains faster than int.Parse(ReadOnlySpan)
    /// due to bounds-check optimization and inline-friendly control flow.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static int IntParse(this ReadOnlySpan<char> value)
    {
        // Simple, branch-friendly integer parse optimized for .NET 8/9 JIT.
        int result = 0;
        int i = 0;
        bool negative = false;
        int len = value.Length;
        if (len == 0)
        {
            return 0;
        }

        char c0 = value[0];
        if (c0 == '-')
        {
            negative = true;
            i = 1;
        }

        for (; i < len; i++)
        {
            ref readonly char ch = ref value[i];
            if ((uint)(ch - '0') > 9u)
            {
                break; // non-digit
            }

            result = (result * 10) + (ch - '0');
        }

        return negative ? -result : result;
    }

    /// <summary>
    /// Reads until the element with the specified local name is found.
    /// </summary>
    /// <param name="reader"></param>
    /// <param name="localName"></param>
    /// <returns></returns>
    public static async Task<bool> ReadToFollowingAsync(this System.Xml.XmlReader reader, string localName)
    {
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (reader.NodeType == System.Xml.XmlNodeType.Element
                && reader.LocalName == localName)
            {
                return true;
            }
        }
        return false;
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

}
