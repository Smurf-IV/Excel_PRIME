using System;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

using ExcelPRIME.Implementation;

#if NET8_0_OR_GREATER
#endif

namespace ExcelPRIME;

#pragma warning disable CA2225 // Implement To### as partner to operator overloads. -> Already exists due to As### properties.

/// <summary>
/// Represents a strongly-typed cell value with custom ToString conversion.
/// Supports zero-allocation formatting on .NET 8+ via ISpanFormattable.
/// </summary>
[StructLayout(LayoutKind.Explicit)]
public readonly struct CellValue : IEquatable<CellValue>, ISpanFormattable, IFormattable
{
    [FieldOffset(0)] private readonly string? _strValue;
    [FieldOffset(8)] private readonly bool _boolValue;
    [FieldOffset(8)] private readonly double _doubleValue;
    [FieldOffset(8)] private readonly DateTime _dateTimeValue;
    [FieldOffset(16)] private readonly CellValueType _type;
    [FieldOffset(20)] private readonly int _iStyleRef; // specifies the identifier of the "cell Formatting", i.e. number of decimals etc.

    private enum CellValueType
    {
        Unknown,
        Numeric,
        String,
        Bool,
        Error,
        DateTime,
        IsDBNull
    }

    // Micro-optimization: Cache frequently allocated strings
    private static readonly CultureInfo InvariantCultureCache = CultureInfo.InvariantCulture;

    // Remove AggressiveOptimization from Constructors
    internal CellValue(string? strValue, int iStyleRef)
    {
        // TODO: iStyleRef might make the string conversion different in future
        _boolValue = default;
        _doubleValue = default;
        _dateTimeValue = default;
        _strValue = strValue ?? string.Empty;
        _type = CellValueType.String;
        _iStyleRef = iStyleRef;
    }

    // Remove AggressiveOptimization from Constructors
    internal CellValue(bool boolValue)
    {
        _strValue = default;
        _doubleValue = default;
        _dateTimeValue = default;
        _boolValue = boolValue;
        _type = CellValueType.Bool;
        _iStyleRef = 0;
    }

    // Remove AggressiveOptimization from Constructors
    internal CellValue(double doubleValue, int iStyleRef)
    {
        _strValue = default;
        _boolValue = default;
        _dateTimeValue = default;
        _doubleValue = doubleValue;
        _type = CellValueType.Numeric;
        _iStyleRef = iStyleRef;
    }

    // Remove AggressiveOptimization from Constructors
    internal CellValue(DateTime dateTimeValue, int iStyleRef)
    {
        _strValue = default;
        _boolValue = default;
        _doubleValue = default;
        _dateTimeValue = dateTimeValue;
        _type = CellValueType.DateTime;
        _iStyleRef = iStyleRef;
    }

    // Remove AggressiveOptimization from Constructors
    internal CellValue(ExcelErrorCode errorCodeValue)
    {
        _strValue = default;
        _boolValue = default;
        _dateTimeValue = default;
        _doubleValue = (int)errorCodeValue;
        _type = CellValueType.Error;
        _iStyleRef = 0;
    }

    internal CellValue(DBNull _/*isDBNull*/, int iStyleRef)
    {
        _strValue = default;
        _boolValue = default;
        _doubleValue = default;
        _dateTimeValue = default;
        _iStyleRef = iStyleRef;
        _type = CellValueType.IsDBNull;
    }

    /// <summary>
    /// Converts the cell value to its string representation.
    /// </summary>
    /// <returns>
    /// A string representation of the cell value, or <c>null</c> if the value is <c>null</c>.
    /// If the value is a string, it is returned as-is. If the value implements <see cref="IFormattable"/>,
    /// it is formatted using the invariant culture. Otherwise, the default <see cref="object.ToString"/> 
    /// implementation is used.
    /// </returns>
    // Remove AggressiveOptimization from Constructors
    public override string? ToString()
    {
        if (_strValue != null || _type == CellValueType.String)
        {
            return _strValue;
        }

        return ToString_Slow();
    }

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private string? ToString_Slow() =>
        _type switch
        {
            CellValueType.Bool => _boolValue ? bool.TrueString : bool.FalseString,
            // Micro-optimization: Use cached CultureInfo instead of property access
            CellValueType.Numeric => _doubleValue.ToString(InvariantCultureCache),
            CellValueType.DateTime => _dateTimeValue.ToString(InvariantCultureCache),
            CellValueType.Error => ((ExcelErrorCode)_doubleValue).ToString(),
            CellValueType.IsDBNull => DBNull.Value.ToString(InvariantCultureCache),
            _ => null
        };

    /// <summary>
    /// Appends the cell value to the provided StringBuilder without allocating an intermediate string.
    /// This is more efficient than calling ToString() when building formatted output.
    /// </summary>
    /// <param name="builder">The StringBuilder to append to.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="builder"/> is null.</exception>
    internal void AppendTo(StringBuilder builder)
    {
        ArgumentNullException.ThrowIfNull(builder);
#if NET8_0_OR_GREATER
        // Try zero-allocation formatting using stackalloc buffer
        Span<char> buffer = stackalloc char[64]; // Adjust size as needed
        if (TryFormat(buffer, out int charsWritten, default, null))
        {
            builder.Append(buffer.Slice(0, charsWritten));
            return;
        }
        // Fallback to string-based formatting if buffer is too small
        builder.Append(ToString());
#else
        // Fast path: string value, no conversion
        if (_strValue != null || _type == CellValueType.String)
        {
            if (_strValue != null)
                builder.Append(_strValue);
            return;
        }
        // Dispatch to type-specific formatting
        AppendTo_Slow(builder);
#endif
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private void AppendTo_Slow(StringBuilder builder)
    {
        switch (_type)
        {
            case CellValueType.Bool:
                builder.Append(_boolValue ? bool.TrueString : bool.FalseString);
                break;

            case CellValueType.Numeric:
                builder.Append(_doubleValue);
                break;

            case CellValueType.DateTime:
                builder.Append(_dateTimeValue);
                break;

            case CellValueType.Error:
                builder.Append((ExcelErrorCode)_doubleValue);
                break;

            case CellValueType.IsDBNull:
                builder.Append(DBNull.Value);
                break;
        }
    }

    /// <summary>
    /// Gets the raw "Boxed" value of the cell.
    /// </summary>
    public object? BoxedValue =>
        _type switch
        {
            CellValueType.Unknown => null,
            CellValueType.Bool => _boolValue,
            CellValueType.Numeric => _doubleValue,
            CellValueType.DateTime => _dateTimeValue,
            CellValueType.Error => (ExcelErrorCode)_doubleValue,
            CellValueType.String => _strValue,
            CellValueType.IsDBNull => DBNull.Value,
            _ => null
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="DateTime"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="DateTime"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public DateTime AsDateTime =>
        // Simplified without branches for common case
        _type == CellValueType.DateTime
            ? _dateTimeValue
            : AsDateTime_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private DateTime AsDateTime_Slow() =>
        _type switch
        {
            CellValueType.Numeric => DateTime.FromOADate(_doubleValue),
            _ => double.TryParse(_strValue, out double val)
                ? // Excel stores the DateTime as a double OADate
                DateTime.FromOADate(val)
                : DateTime.Parse(ToString()!, CultureInfo.InvariantCulture)
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="DateOnly"/> object, if possible.
    /// </summary>
    public DateOnly AsDateOnly =>
        DateOnly.FromDateTime(AsDateTime);

    /// <summary>
    /// Gets the value of the cell as a <see cref="TimeOnly"/> object, if possible.
    /// </summary>
    public TimeOnly AsTimeOnly =>
        TimeOnly.FromDateTime(AsDateTime);

    /// <summary>
    /// Gets the value of the cell as a <see cref="bool"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="bool"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public bool AsBoolean =>
        // Simplified without branches for common case
        _type == CellValueType.Bool
            ? _boolValue
            : AsBoolean_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private bool AsBoolean_Slow() =>
        _type switch
        {
            CellValueType.DateTime => _dateTimeValue.Ticks != 0,
            CellValueType.Error => (ExcelErrorCode)_doubleValue != ExcelErrorCode.Null,
            CellValueType.Numeric => _doubleValue != 0,
            CellValueType.IsDBNull => false,
            _ => int.TryParse(_strValue, out int val) ? val != 0 : Convert.ToBoolean(_strValue!)
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="Int32"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Int32"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public int AsInt32 =>
        // Simplified without branches for common case
        _type == CellValueType.Numeric
            ? (int)_doubleValue
            : AsInt32_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private int AsInt32_Slow() =>
        _type switch
        {
            CellValueType.DateTime => (int)_dateTimeValue.Ticks,
            CellValueType.Bool => _boolValue ? 1 : 0,
            CellValueType.Error => (int)_doubleValue,
            CellValueType.IsDBNull => 0,
            _ => int.Parse(_strValue!, NumberStyles.Integer, CultureInfo.InvariantCulture)
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="Int64"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Int64"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public long AsInt64 =>
        _type == CellValueType.Numeric
            ? (long)_doubleValue
            : AsInt64_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private long AsInt64_Slow() =>
        _type switch
        {
            CellValueType.DateTime => _dateTimeValue.Ticks,
            CellValueType.Bool => _boolValue ? 1 : 0,
            CellValueType.Error => (long)_doubleValue,
            CellValueType.IsDBNull => 0L,
            _ => long.Parse(_strValue!, NumberStyles.Integer, CultureInfo.InvariantCulture)
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="double"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="double"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public double AsDouble =>
        // Simplified without branches for common case
        _type == CellValueType.Numeric
            ? _doubleValue
            : AsDouble_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private double AsDouble_Slow() =>
        _type switch
        {
            CellValueType.DateTime => _dateTimeValue.ToOADate(),
            CellValueType.Bool => _boolValue ? 1 : 0,
            CellValueType.Error => _doubleValue,
            CellValueType.IsDBNull => 0.0,
            _ => double.Parse(_strValue!, NumberStyles.Float, CultureInfo.InvariantCulture)
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="Decimal"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Decimal"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public decimal AsDecimal =>
        // Simplified without branches for common case
        _type == CellValueType.Numeric
            ? (decimal)_doubleValue
            : AsDecimal_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private decimal AsDecimal_Slow() =>
        _type switch
        {
            CellValueType.DateTime => (decimal)_dateTimeValue.ToOADate(),
            CellValueType.Bool => _boolValue ? 1 : 0,
            CellValueType.Error => (decimal)_doubleValue,
            CellValueType.IsDBNull => 0m,
            _ => decimal.Parse(_strValue!, NumberStyles.Currency, CultureInfo.InvariantCulture)
        };

    #region Styled Formatters

    /// <summary>
    /// Returns the cell value as a string formatted according to the specified styles dictionary, if available.
    /// </summary>
    /// <returns>
    /// A string representation of the cell value formatted according to the cell's style, 
    /// or the default string representation if the style is not found.
    /// </returns>
    public string? ToStyledString()
    {
        if (_iStyleRef < 0
            || !Ecma376StandardProvider.TryGetFormat(_iStyleRef, out string? formatCode, out FormattingType type)
            || string.IsNullOrWhiteSpace(formatCode)
            || type == FormattingType.General
            )
        {
            return ToString();
        }
        return FormatValueWithStyle(formatCode, type);
    }

    /// <summary>
    /// Formats the cell value according to the provided style.
    /// </summary>
    private string? FormatValueWithStyle(string formatCode, FormattingType type) =>
        _type switch
        {
            CellValueType.Numeric => FormatNumericWithNumberFormat(_doubleValue, formatCode, type),
            CellValueType.DateTime => FormatDateTimeWithNumberFormat(_dateTimeValue, formatCode, type),
            CellValueType.Bool => _boolValue ? bool.TrueString : bool.FalseString,
            _ => ToString() // TODO: What happens if the formatCode is applied to a string type ?
        };

    /// <summary>
    /// Formats a numeric value according to an Excel number format code.
    /// Handles all FormattingType.Number styles including:
    /// - Regular numbers (0, 0.00, #,##0, #,##0.00)
    /// - Scientific notation (0.00E+00, ##0.0E0, etc.)
    /// - Percentages (0%, 0.0%, 0.00%)
    /// - Fractions (# ?/?, # ??/??)
    /// - International formats and variants
    /// </summary>
    private static string FormatNumericWithNumberFormat(double value, string formatCode, FormattingType type)
    {
        // Ensure we're handling FormattingType.Number
        if (type != FormattingType.Number)
        {
            return value.ToString(InvariantCultureCache);
        }

        return formatCode switch
        {
            // General and text formats
            "General" => value.ToString(InvariantCultureCache),
            "@" => value.ToString(InvariantCultureCache),

            // Basic integer formats
            "0" => Math.Round(value).ToString(InvariantCultureCache),

            // Decimal formats
            "0.00" => value.ToString("F2", InvariantCultureCache),
            "0.0" => value.ToString("F1", InvariantCultureCache),

            // Thousand separator formats
            "#,##0" => Math.Round(value).ToString("N0", InvariantCultureCache),
            "#,##0.0" => value.ToString("N1", InvariantCultureCache),
            "#,##0.00" => value.ToString("N2", InvariantCultureCache),

            // With brackets (negative in parentheses)
            "#,##0;(#,##0)" => FormatNumberWithNegativeParentheses(value, "N0"),
            "#,##0.0;(#,##0.0)" => FormatNumberWithNegativeParentheses(value, "N1"),
            "#,##0.00;(#,##0.00)" => FormatNumberWithNegativeParentheses(value, "N2"),

            // With red color indicator for negatives
            "#,##0;[Red](#,##0)" => FormatNumberWithNegativeParentheses(value, "N0"),
            "#,##0.00;[Red](#,##0.00)" => FormatNumberWithNegativeParentheses(value, "N2"),

            // Currency/Accounting formats (non-currency symbol versions)
            "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)" => FormatAccountingNumber(value, 0),
            "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)" => FormatAccountingNumber(value, 2),

            // Percentage formats
            "0%" => (value * 100).ToString("F0", InvariantCultureCache) + "%",
            "0.0%" => (value * 100).ToString("F1", InvariantCultureCache) + "%",
            "0.00%" => (value * 100).ToString("F2", InvariantCultureCache) + "%",

            // Scientific notation
            "0.00E+00" => value.ToString("E2", InvariantCultureCache),
            "0.00E+0" => value.ToString("E2", InvariantCultureCache),
            "0.00E0" => value.ToString("E2", InvariantCultureCache),
            "##0.0E0" => value.ToString("E1", InvariantCultureCache),
            "##0.0E+0" => value.ToString("E1", InvariantCultureCache),
            "##0.0E+00" => value.ToString("E1", InvariantCultureCache),

            // Fraction formats
            "# ?/?" => FormatFraction(value, 1),
            "# ??/??" => FormatFraction(value, 2),

            // CJK formats (treated as numbers)
            "[DBNum1][$-804]0" => value.ToString("F0", InvariantCultureCache),
            "[DBNum1][$-804]0.00" => value.ToString("F2", InvariantCultureCache),
            "[DBNum4][$-804]0" => value.ToString("F0", InvariantCultureCache),

            // Default: return as double with invariant culture
            _ => FormatCustomNumber(value, formatCode)
        };
    }

    /// <summary>
    /// Formats a number with negative values in parentheses.
    /// </summary>
    private static string FormatNumberWithNegativeParentheses(double value, string format)
    {
        if (value < 0)
        {
            return string.Concat("(", Math.Abs(value).ToString(format, CultureInfo.InvariantCulture), ")");
        }
        return value.ToString(format, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Formats a number in accounting style with alignment spacing.
    /// </summary>
    private static string FormatAccountingNumber(double value, int decimals)
    {
        string format = decimals switch
        {
            0 => "N0",
            1 => "N1",
            2 => "N2",
            _ => "N" + decimals
        };

        if (value < 0)
        {
            return string.Concat("(", Math.Abs(value).ToString(format, CultureInfo.InvariantCulture), ")");
        }

        // Add leading space for alignment
        return string.Concat(" ", value.ToString(format, CultureInfo.InvariantCulture), " ");
    }

    /// <summary>
    /// Formats a decimal number as a fraction.
    /// </summary>
    private static string FormatFraction(double value, int maxDigits)
    {
        // Get the integer and fractional parts
        double intPart = Math.Truncate(value);
        double fracPart = Math.Abs(value - intPart);

        // Find the best rational approximation
        int denominator = maxDigits switch
        {
            1 => 8,  // Single digit: /1 through /9, common: /8
            2 => 16, // Double digit: /16 is common for 1/16
            _ => 8
        };

        // Try to find a simpler fraction
        int numerator = (int)Math.Round(fracPart * denominator);

        // Simplify fraction if needed
        int gcd = GreatestCommonDivisor(numerator, denominator);
        numerator /= gcd;
        denominator /= gcd;

        if (numerator == 0)
        {
            return Math.Round(value).ToString(CultureInfo.InvariantCulture);
        }

        string result = string.Concat(numerator, "/", denominator);

        if (intPart != 0)
        {
            result = string.Concat(Math.Truncate(intPart).ToString(CultureInfo.InvariantCulture), " ", result);
        }

        return result;
    }

    /// <summary>
    /// Calculates the greatest common divisor using Euclidean algorithm.
    /// </summary>
    private static int GreatestCommonDivisor(int a, int b)
    {
        while (b != 0)
        {
            int temp = b;
            b = a % b;
            a = temp;
        }
        return a;
    }

    /// <summary>
    /// Formats a numeric value using a custom format code pattern.
    /// </summary>
    private static string FormatCustomNumber(double value, string formatCode)
    {
        var span = formatCode.AsSpan();
        // Check for percentage format in the code
        if (span.Contains('%'))
        {
            // Count decimal places in the format
            int decimalPlaces = 0;
            int dotIndex = span.IndexOf('.');
            if (dotIndex >= 0)
            {
                var afterDot = span.Slice(dotIndex + 1);
                foreach (char c in afterDot)
                {
                    if (c == '0')
                        decimalPlaces++;
                    else
                        break;
                }
            }

            return (value * 100).ToString("F" + decimalPlaces, CultureInfo.InvariantCulture) + "%";
        }

        // Check for scientific notation
        if (span.ContainsAny("Ee"))
        {
            int eIndex = Math.Max(span.IndexOf('E'), span.IndexOf('e'));
            int decimalPlaces = 2; // default
            if (eIndex > 0)
            {
                var beforeE = span.Slice(0, eIndex);
                int dotIndex = beforeE.LastIndexOf('.');
                if (dotIndex >= 0)
                {
                    decimalPlaces = beforeE.Length - dotIndex - 1;
                }
            }
            return value.ToString("E" + decimalPlaces, CultureInfo.InvariantCulture);
        }

        // Count decimal places from format code
        int dotPos = span.IndexOf('.');
        int decimals = 0;
        if (dotPos >= 0)
        {
            var afterDot = span.Slice(dotPos + 1);
            foreach (char c in afterDot)
            {
                if (c == '0' || c == '#')
                    decimals++;
            }
        }

        // Check if thousands separator is present
        if (span.Contains(','))
        {
            return value.ToString("N" + decimals, CultureInfo.InvariantCulture);
        }

        // Default fixed-point format
        if (decimals > 0)
        {
            return value.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        return Math.Round(value).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Formats a DateTime value according to an Excel number format code.
    /// Handles all FormattingType.TimeOnly, DateTime, and DateOnly styles including:
    /// - Time-only formats (h:mm, h:mm:ss, mm:ss, [h]:mm:ss)
    /// - Date-only formats (mm/dd/yyyy, d-mmm-yy, yyyy-mm-dd, etc.)
    /// - Combined date/time formats (m/d/yy h:mm, dd/mm/yyyy h:mm:ss, etc.)
    /// - International variants (German, Chinese, etc.)
    /// - Millisecond precision formats (hh:mm:ss.000)
    /// </summary>
    private static string FormatDateTimeWithNumberFormat(DateTime value, string formatCode, FormattingType type) =>
        formatCode switch
        {
            // ============ TimeOnly Formats ============
            "mm:ss" => value.ToString("mm:ss", CultureInfo.InvariantCulture),
            "mm:ss.0" => value.ToString("mm:ss.f", CultureInfo.InvariantCulture),
            "[h]:mm:ss" => FormatElapsedTime(value),

            "h:mm AM/PM" => value.ToString("h:mm tt", CultureInfo.InvariantCulture),
            "h:mm:ss AM/PM" => value.ToString("h:mm:ss tt", CultureInfo.InvariantCulture),

            "h:mm" => value.ToString("h:mm", CultureInfo.InvariantCulture),
            "h:mm:ss" => value.ToString("h:mm:ss", CultureInfo.InvariantCulture),

            "hh:mm:ss" => value.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
            "hh:mm:ss.000" => value.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture),
            "h:mm:ss.00" => value.ToString("h:mm:ss.ff", CultureInfo.InvariantCulture),

            // Locale-specific time formats (US)
            "[$-409]h:mm AM/PM" => value.ToString("h:mm tt", CultureInfo.InvariantCulture),
            "[$-409]h:mm:ss AM/PM" => value.ToString("h:mm:ss tt", CultureInfo.InvariantCulture),
            "[$-409]h:mm" => value.ToString("h:mm", CultureInfo.InvariantCulture),
            "[$-409]h:mm:ss" => value.ToString("h:mm:ss", CultureInfo.InvariantCulture),

            // ============ DateOnly Formats ============
            "mm/dd/yyyy" => value.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture),
            "m/d/yy" => value.ToString("M/d/yy", CultureInfo.InvariantCulture),
            "mm/dd/yy" => value.ToString("MM/dd/yy", CultureInfo.InvariantCulture),

            "d-mmm-yy" => value.ToString("d-MMM-yy", CultureInfo.InvariantCulture),
            "d-mmm" => value.ToString("d-MMM", CultureInfo.InvariantCulture),
            "mmm-yy" => value.ToString("MMM-yy", CultureInfo.InvariantCulture),

            "d/m/yy" => value.ToString("d/M/yy", CultureInfo.InvariantCulture),
            "d.m.yy" => value.ToString("d.M.yy", CultureInfo.InvariantCulture),
            "d.m.yyyy" => value.ToString("d.M.yyyy", CultureInfo.InvariantCulture),

            "yyyy-m-d" => value.ToString("yyyy-M-d", CultureInfo.InvariantCulture),
            "yyyy-mm-dd" => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),

            "dd-mmm-yyyy" => value.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture),
            "dd/mmm/yyyy" => value.ToString("dd/MMM/yyyy", CultureInfo.InvariantCulture),

            "dd-mm-yy" => value.ToString("dd-MM-yy", CultureInfo.InvariantCulture),
            "dd-mm-yyyy" => value.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture),

            "dd MMMM yyyy" => value.ToString("dd MMMM yyyy", CultureInfo.InvariantCulture),
            "d. MMMM yyyy" => value.ToString("d. MMMM yyyy", CultureInfo.InvariantCulture),

            "d MMM yy" => value.ToString("d MMM yy", CultureInfo.InvariantCulture),
            "d MMMM yy" => value.ToString("d MMMM yy", CultureInfo.InvariantCulture),

            "mm-dd" => value.ToString("MM-dd", CultureInfo.InvariantCulture),
            "mm-dd-yy" => value.ToString("MM-dd-yy", CultureInfo.InvariantCulture),
            "mm-dd-yyyy" => value.ToString("MM-dd-yyyy", CultureInfo.InvariantCulture),

            "mmmm d, yyyy" => value.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture),
            "d MMMM, yyyy" => value.ToString("d MMMM, yyyy", CultureInfo.InvariantCulture),

            "dd-mmm" => value.ToString("dd-MMM", CultureInfo.InvariantCulture),

            "ddd, mmmm dd, yyyy" => value.ToString("ddd, MMMM dd, yyyy", CultureInfo.InvariantCulture),

            // CJK and Japanese formats
            "g/m/d" => value.ToString("g/M/d", CultureInfo.InvariantCulture),
            "ge.m.d" => value.ToString("ge.M.d", CultureInfo.InvariantCulture),
            "gg" => value.ToString("gg", CultureInfo.InvariantCulture),
            "ggg" => value.ToString("ggg", CultureInfo.InvariantCulture),

            // Locale-specific date formats (US)
            "[$-409]M/d/yy" => value.ToString("M/d/yy", CultureInfo.InvariantCulture),
            "[$-409]d-mmm-yy" => value.ToString("d-MMM-yy", CultureInfo.InvariantCulture),
            "[$-409]d-mmm" => value.ToString("d-MMM", CultureInfo.InvariantCulture),
            "[$-409]mmm-yy" => value.ToString("MMM-yy", CultureInfo.InvariantCulture),

            // German date format
            "d. mmm. yyyy" => value.ToString("d. MMM. yyyy", CultureInfo.InvariantCulture),
            "dddd, d. mmmm yyyy" => value.ToString("dddd, d. MMMM yyyy", CultureInfo.InvariantCulture),

            // ============ DateTime (Combined) Formats ============
            "m/d/yy h:mm" => value.ToString("M/d/yy h:mm", CultureInfo.InvariantCulture),
            "m/d/yy h:mm:ss" => value.ToString("M/d/yy h:mm:ss", CultureInfo.InvariantCulture),

            "d/m/yy h:mm" => value.ToString("d/M/yy h:mm", CultureInfo.InvariantCulture),
            "d/m/yy h:mm:ss" => value.ToString("d/M/yy h:mm:ss", CultureInfo.InvariantCulture),

            "d/m/yyyy h:mm" => value.ToString("d/M/yyyy h:mm", CultureInfo.InvariantCulture),
            "d/m/yyyy h:mm:ss" => value.ToString("d/M/yyyy h:mm:ss", CultureInfo.InvariantCulture),

            "yyyy-m-d h:mm:ss" => value.ToString("yyyy-M-d h:mm:ss", CultureInfo.InvariantCulture),

            "mm/dd/yyyy h:mm:ss" => value.ToString("MM/dd/yyyy h:mm:ss", CultureInfo.InvariantCulture),
            "dd/mm/yyyy h:mm:ss" => value.ToString("dd/MM/yyyy h:mm:ss", CultureInfo.InvariantCulture),

            "yyyy-mm-dd hh:mm:ss" => value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            "yyyy-mm-dd'T'hh:mm:ss" => value.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture),

            "mmmm d, yyyy h:mm:ss" => value.ToString("MMMM d, yyyy h:mm:ss", CultureInfo.InvariantCulture),

            // Locale-specific datetime formats (US)
            "[$-409]m/d/yy h:mm" => value.ToString("M/d/yy h:mm", CultureInfo.InvariantCulture),

            // ============ Default/Fallback ============
            _ => FormatCustomDateTime(value, formatCode, type)
        };

    /// <summary>
    /// Formats elapsed time in [h]:mm:ss format (hours can exceed 24).
    /// </summary>
    private static string FormatElapsedTime(DateTime value)
    {
        TimeOnly timeOnly = TimeOnly.FromDateTime(value);
        int totalHours = timeOnly.Hour;
        int minutes = value.Minute;
        int seconds = value.Second;
        return $"{totalHours}:{minutes:D2}:{seconds:D2}";
    }

    /// <summary>
    /// Formats a DateTime value using a custom format code pattern.
    /// </summary>
    private static string FormatCustomDateTime(DateTime value, string formatCode, FormattingType type)
    {
        // Handle formats with color indicators (e.g., "[Red]mm:ss")
        if (formatCode.Contains("[Red]"))
        {
            string cleanFormat = formatCode.Replace("[Red]", "");
            return FormatCustomDateTime(value, cleanFormat, type);
        }

        // Try to convert Excel format codes to .NET format codes
        string netFormat = ConvertExcelDateTimeFormatToNet(formatCode);

        try
        {
            return value.ToString(netFormat, CultureInfo.InvariantCulture);
        }
        catch
        {
            // Fallback based on FormattingType
            return type switch
            {
                FormattingType.TimeOnly => value.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
                FormattingType.DateOnly => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                FormattingType.DateTime => value.ToString("s", CultureInfo.InvariantCulture),
                _ => value.ToString(CultureInfo.InvariantCulture)
            };
        }
    }

    /// <summary>
    /// Converts Excel date/time format codes to .NET DateTime format strings.
    /// Intelligently handles "mm" which can be either months or minutes depending on context.
    /// </summary>
    private static string ConvertExcelDateTimeFormatToNet(string excelFormat)
    {
        // Build a mapping for common conversions
        string result = excelFormat;

        // Handle full and abbreviated month names first (these are unambiguous)
        result = result.Replace("mmmm", "MMMM"); // Full month name
        result = result.Replace("mmm", "MMM");   // Abbreviated month name

        // Now handle "mm" carefully - only replace when it's a month, not minutes
        // "mm" is a month when:
        // - It's followed by "/" or "-" and preceded/followed by day or year (e.g., "mm/dd", "d-mm")
        // - It's part of date format
        // "mm" is minutes when:
        // - It's followed by ":" (e.g., "h:mm", "mm:ss")
        // - It's in a time-only context

        // Safer approach: replace "mm" only when NOT followed by colon
        // Build the result character by character to avoid replacing "mm" in time contexts
        StringBuilder sb = new StringBuilder(result.Length);
        int i = 0;
        while (i < result.Length)
        {
            // Look for "mm" pattern
            if (i < result.Length - 1 && result[i] == 'm' && result[i + 1] == 'm')
            {
                // Check if it's followed by a colon (indicating minutes)
                if (i + 2 < result.Length && result[i + 2] == ':')
                {
                    // This is minutes - keep as "mm"
                    sb.Append("mm");
                    i += 2;
                }
                else
                {
                    // Check if preceded by digit or date separator, or followed by date pattern
                    bool isProbablyMonth = false;

                    // If preceded by a digit or date separator
                    if (i > 0 && (char.IsDigit(result[i - 1]) || result[i - 1] == '/' || result[i - 1] == '-' || result[i - 1] == '.'))
                    {
                        isProbablyMonth = true;
                    }

                    // If followed by digit, date separator, or date character
                    if (i + 2 < result.Length)
                    {
                        char next = result[i + 2];
                        if (char.IsDigit(next) || next == '/' || next == '-' || next == '.' || next == 'd' || next == 'y' || next == 'D' || next == 'Y')
                        {
                            isProbablyMonth = true;
                        }
                    }

                    if (isProbablyMonth)
                    {
                        sb.Append("MM");
                    }
                    else
                    {
                        sb.Append("mm");
                    }
                    i += 2;
                }
            }
            else
            {
                sb.Append(result[i]);
                i++;
            }
        }
        result = sb.ToString();

        // Year patterns
        //result = result.Replace("yyyy", "yyyy"); // Four-digit year
        //result = result.Replace("yy", "yy");     // Two-digit year

        // Time patterns
        //result = result.Replace("H", "H");       // One or two-digit hour (24-hour)
        //result = result.Replace("ss", "ss");     // Two-digit seconds
        //result = result.Replace("ff", "ff");     // Two-digit milliseconds
        //result = result.Replace("fff", "fff");   // Three-digit milliseconds

        // AM/PM indicators
        result = result.Replace("AM/PM", "tt");     // AM/PM
        result = result.Replace("A/P", "t");        // A/P

        return result;
    }

    #endregion

    /// <summary>
    /// Determines whether the specified object is equal to the current<see cref="CellValue"/> instance.
    /// </summary>
    /// <param name = "obj" > The object to compare with the current <see cref = "CellValue" /> instance.</param >
    /// <returns>
    /// <c> true </c> if the specified object is a<see cref = "CellValue" /> and has the same type and value as the current instance; otherwise, <c>false</c>.
    /// </returns>
    public override bool Equals(object? obj) => obj is CellValue other && Equals(other);

    /// <summary>
    /// Returns the hash code for the current<see cref="CellValue"/> instance.
    /// </summary>
    /// <returns>
    /// A 32-bit signed integer hash code that represents the current<see cref = "CellValue" />.
    /// </returns >
    /// <remarks>
    /// The hash code is computed based on the type of the cell value and its associated data,
    /// ensuring that equal<see cref = "CellValue" /> instances produce the same hash code.
    /// </remarks>
    public override int GetHashCode() => HashCode.Combine(_type, _boolValue, _doubleValue, _dateTimeValue, _strValue, _iStyleRef);

    /// <summary>
    /// Determines whether the current<see cref="CellValue"/> instance is equal to another<see cref = "CellValue" /> instance.
    /// </summary >
    /// <param name= "other" > The <see cref= "CellValue" /> instance to compare with the current instance.</param>
    /// <returns>
    /// <c>true</c> if the current instance and the<paramref name = "other" /> instance are equal; otherwise, <c>false</c>.
    /// </returns>
    /// <remarks>
    /// Equality is determined based on the<see cref="CellType"/> and the value associated with it.
    /// For example, numeric values are compared numerically, strings are compared using string equality,
    /// and dates are compared using date-time equality.
    /// </remarks>
    public bool Equals(CellValue other)
    {
        if (_type != other._type)
        {
            return false;
        }

        return _type switch
        {
            CellValueType.Bool => _boolValue == other._boolValue,
            CellValueType.Numeric => _doubleValue == other._doubleValue,
            CellValueType.DateTime => _dateTimeValue == other._dateTimeValue,
            //CellValueType.String => _strValue == other._strValue,
            CellValueType.IsDBNull => true,
            _ => _strValue == other._strValue
        };
    }

    /// <summary>
    /// Determines whether two<see cref = "CellValue" /> instances are equal.
    /// </summary>
    /// <param name = "left" > The first<see cref = "CellValue" /> to compare.</param>
    /// <param name = "right" > The second<see cref = "CellValue" /> to compare.</param>
    /// <returns>
    /// <c>true</c> if the specified <see cref = "CellValue" /> instances are equal; otherwise, <c>false</c>.
    /// </returns>
    public static bool operator ==(CellValue left, CellValue right) => left.Equals(right);

    /// <summary>
    /// Determines whether two<see cref = "CellValue" /> instances are not equal.
    /// </summary>
    /// <param name = "left" > The first<see cref = "CellValue" /> to compare.</param>
    /// <param name = "right" > The second<see cref = "CellValue" /> to compare.</param>
    /// <returns>
    /// <c>true</c> if the specified<see cref="CellValue"/> instances are not equal; otherwise, <c>false</c>.
    /// </returns>
    public static bool operator !=(CellValue left, CellValue right) => !left.Equals(right);

    /// <summary>
    /// Attempts to get the value of the cell as a <see cref="DateTime"/> object.
    /// </summary>
    /// <param name="value">
    /// When this method returns, contains the <see cref="DateTime"/> value of the cell if the conversion succeeded, 
    /// or the default value if the conversion failed.
    /// </param>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="DateTime"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryGetDateTime(out DateTime value)
    {
        try
        {
            value = AsDateTime;
            return true;
        }
        catch
        {
            value = default;
            return false;
        }
    }

    /// <summary>
    /// Gets the value of the cell as a <see cref="DateOnly"/> object, if possible.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="DateOnly"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryDateOnly(out DateOnly value)
    {
        try
        {
            value = DateOnly.FromDateTime(AsDateTime);
            return true;
        }
        catch
        {
            value = default;
            return false;
        }
    }

    /// <summary>
    /// Gets the value of the cell as a <see cref="TimeOnly"/> object, if possible.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="TimeOnly"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryTimeOnly(out TimeOnly value)
    {
        try
        {
            value = TimeOnly.FromDateTime(AsDateTime);
            return true;
        }
        catch
        {
            value = default;
            return false;
        }
    }

    /// <summary>
    /// Attempts to get the value of the cell as a <see cref="bool"/> object.
    /// </summary>
    /// <param name="value">
    /// When this method returns, contains the <see cref="bool"/> value of the cell if the conversion succeeded, 
    /// or the default value if the conversion failed.
    /// </param>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="bool"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryGetBoolean(out bool value)
    {
        try
        {
            value = AsBoolean;
            return true;
        }
        catch
        {
            value = false;
            return false;
        }
    }

    /// <summary>
    /// Attempts to get the value of the cell as a <see cref="int"/> object.
    /// </summary>
    /// <param name="value">
    /// When this method returns, contains the <see cref="int"/> value of the cell if the conversion succeeded, 
    /// or the default value if the conversion failed.
    /// </param>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="int"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryGetInt32(out int value)
    {
        try
        {
            value = AsInt32;
            return true;
        }
        catch
        {
            value = 0;
            return false;
        }
    }

    /// <summary>
    /// Attempts to get the value of the cell as a <see cref="long"/> object.
    /// </summary>
    /// <param name="value">
    /// When this method returns, contains the <see cref="long"/> value of the cell if the conversion succeeded, 
    /// or the default value if the conversion failed.
    /// </param>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="long"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryGetInt64(out long value)
    {
        try
        {
            value = AsInt64;
            return true;
        }
        catch
        {
            value = 0;
            return false;
        }
    }

    /// <summary>
    /// Attempts to get the value of the cell as a <see cref="double"/> object.
    /// </summary>
    /// <param name="value">
    /// When this method returns, contains the <see cref="double"/> value of the cell if the conversion succeeded, 
    /// or the default value if the conversion failed.
    /// </param>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="double"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryGetDouble(out double value)
    {
        try
        {
            value = AsDouble;
            return true;
        }
        catch
        {
            value = 0;
            return false;
        }
    }

    /// <summary>
    /// Attempts to get the value of the cell as a <see cref="decimal"/> object.
    /// </summary>
    /// <param name="value">
    /// When this method returns, contains the <see cref="decimal"/> value of the cell if the conversion succeeded, 
    /// or the default value if the conversion failed.
    /// </param>
    /// <returns>
    /// <c>true</c> if the cell value was successfully converted to a <see cref="decimal"/>; otherwise, <c>false</c>.
    /// </returns>
    public bool TryGetDecimal(out decimal value)
    {
        try
        {
            value = AsDecimal;
            return true;
        }
        catch
        {
            value = 0;
            return false;
        }
    }

    #region Implicit Operators

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="string"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A string representation of the cell value, or <c>null</c> if the value is <c>null</c>.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator string?(CellValue value) => value.ToString();

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="bool"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A boolean representation of the cell value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator bool(CellValue value) => value.AsBoolean;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="int"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A 32-bit integer representation of the cell value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator int(CellValue value) => value.AsInt32;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="long"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A 64-bit integer representation of the cell value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator long(CellValue value) => value.AsInt64;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="double"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A double representation of the cell value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator double(CellValue value) => value.AsDouble;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="decimal"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A decimal representation of the cell value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator decimal(CellValue value) => value.AsDecimal;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="DateTime"/>.
    /// </summary>
    /// <param name="value">The <see cref="CellValue"/> to convert.</param>
    /// <returns>A DateTime representation of the cell value.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator DateTime(CellValue value) => value.AsDateTime;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="DateOnly"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator DateOnly(CellValue value) => value.AsDateOnly;

    /// <summary>
    /// Implicitly converts a <see cref="CellValue"/> to a <see cref="TimeOnly"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static implicit operator TimeOnly(CellValue value) => value.AsTimeOnly;

    #endregion
#if NET8_0_OR_GREATER
    #region ISpanFormattable Implementation

    /// <summary>
    /// Tries to format the cell value into the provided character span without allocations.
    /// This zero-allocation method is available on .NET 8+.
    /// </summary>
    /// <param name="destination">The span where the formatted value should be written.</param>
    /// <param name="charsWritten">The number of characters written.</param>
    /// <param name="format">Unused; invariant culture is always applied.</param>
    /// <param name="provider">Unused; invariant culture is always applied.</param>
    /// <returns>true if successful; false if destination is too small.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
    {
        // Fast path: string value, no conversion
        if (_strValue != null || _type == CellValueType.String)
        {
            return TryFormatString(destination, out charsWritten);
        }

        // Dispatch to type-specific handling
        return TryFormat_Slow(destination, out charsWritten);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private bool TryFormatString(Span<char> destination, out int charsWritten)
    {
        if (_strValue == null)
        {
            charsWritten = 0;
            return true;
        }

        if (destination.Length < _strValue.Length)
        {
            charsWritten = 0;
            return false;
        }

        _strValue.AsSpan().CopyTo(destination);
        charsWritten = _strValue.Length;
        return true;
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private bool TryFormat_Slow(Span<char> destination, out int charsWritten)
    {
        switch (_type)
        {
            case CellValueType.Numeric:
                return _doubleValue.TryFormat(destination, out charsWritten, default, InvariantCultureCache);

            case CellValueType.Bool:
                {
                    var boolStr = _boolValue ? bool.TrueString : bool.FalseString;
                    if (destination.Length < boolStr.Length)
                    {
                        charsWritten = 0;
                        return false;
                    }
                    boolStr.AsSpan().CopyTo(destination);
                    charsWritten = boolStr.Length;
                    return true;
                }

            case CellValueType.DateTime:
                return _dateTimeValue.TryFormat(destination, out charsWritten, default, InvariantCultureCache);

            case CellValueType.Error:
                return TryFormatError(_doubleValue, destination, out charsWritten);

            case CellValueType.IsDBNull:
                {
                    const string dbNullStr = "DBNull";
                    if (destination.Length < dbNullStr.Length)
                    {
                        charsWritten = 0;
                        return false;
                    }
                    dbNullStr.AsSpan().CopyTo(destination);
                    charsWritten = dbNullStr.Length;
                    return true;
                }

            default:
                charsWritten = 0;
                return true;
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static bool TryFormatError(double value, Span<char> destination, out int charsWritten)
    {
        string errorStr = (ExcelErrorCode)value switch
        {
            ExcelErrorCode.Null => "#NULL!",
            ExcelErrorCode.DivideByZero => "#DIV/0!",
            ExcelErrorCode.Value => "#VALUE!",
            ExcelErrorCode.Reference => "#REF!",
            ExcelErrorCode.Name => "#NAME?",
            ExcelErrorCode.Number => "#NUM!",
            ExcelErrorCode.NotAvailable => "#N/A",
            _ => "Error"
        };

        if (destination.Length < errorStr.Length)
        {
            charsWritten = 0;
            return false;
        }

        errorStr.AsSpan().CopyTo(destination);
        charsWritten = errorStr.Length;
        return true;
    }

    #endregion
#endif


    /// <summary>
    /// Implements IFormattable.ToString for compatibility.
    /// </summary>
    string IFormattable.ToString(string? format, IFormatProvider? formatProvider)
        => ToString() ?? string.Empty;

}
