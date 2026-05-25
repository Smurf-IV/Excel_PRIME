using System.Collections.Concurrent;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;

[assembly: InternalsVisibleTo("Excel_PRIME.Tests")]

namespace ExcelPRIME;

#pragma warning disable CA2225 // Implement To### as partner to operator overloads. -> Already exists due to As### properties.

/// <summary>
/// Represents the type of value stored in a cell.
/// </summary>
internal enum CellValueType : byte
{
    Unknown,
    Decimal,
    Double,
    Int,
    Long,
    String,
    Bool,
    Error,
    DateTime,
    IsDBNull
}

/// <summary>
/// Represents a strongly-typed cell value with custom ToString conversion.
/// Supports zero-allocation formatting on .NET 8+ via ISpanFormattable.
/// </summary>
[StructLayout(LayoutKind.Explicit)]
public class CellValue : IEquatable<CellValue>, ISpanFormattable, IFormattable
{
    #region reduce from 48 bytes to 20 bytes by using explicit layout and overlapping fields
    [FieldOffset(0)] private readonly string? _s; // Stores string values
    /// <summary>
    /// The style reference index.
    /// Specifies the identifier of the "cell Formatting", i.e. number of decimals etc.
    /// </summary>
    [FieldOffset(8)] private readonly short _iStyleRef;
    /// <summary>
    /// The type of the cell value. "byte storage"
    /// </summary>
    [FieldOffset(10)] private readonly CellValueType _type;
    // Offset to nearest 4 byte boundary for better performance of value types
    [FieldOffset(12)] private readonly decimal _d; // Stores value types in a decimal to avoid boxing and precision loss
    [FieldOffset(12)] private readonly DateTime _dt;
    [FieldOffset(12)] private readonly bool _b;
    [FieldOffset(12)] private readonly long _l;
    [FieldOffset(12)] private readonly int _i;
    [FieldOffset(12)] private readonly double _db;
    #endregion


    // Micro-optimization: Cache frequently allocated strings
    private static readonly CultureInfo s_invariantCultureCache = CultureInfo.InvariantCulture;

    private static readonly ConcurrentDictionary<short, CellValue> s_DBNullCache = new();

    /// <summary>
    /// Returns a cached <see cref="CellValue"/> instance representing a DBNull value with the specified style.
    /// </summary>
    /// <param name="iStyleRef">The style reference index.</param>
    /// <returns>A <see cref="CellValue"/> instance.</returns>
    public static CellValue GetDBNull(short iStyleRef)
        => s_DBNullCache.GetOrAdd(iStyleRef, static style => new CellValue(CellValueType.IsDBNull, style));

    internal static CellValue Create(string? strValue, short iStyleRef)
        => new CellValue(strValue ?? string.Empty, CellValueType.String, iStyleRef);

    internal static CellValue Create(bool boolValue)
        => new CellValue(boolValue, CellValueType.Bool, 0);

    internal static CellValue Create(DateTime dateTimeValue, short iStyleRef)
        => new CellValue(dateTimeValue, CellValueType.DateTime, iStyleRef);

    internal static CellValue Create(decimal decimalValue, short iStyleRef)
        => new CellValue(decimalValue, CellValueType.Decimal, iStyleRef);
    internal static CellValue Create(double doubleValue, short iStyleRef)
        => new CellValue(doubleValue, CellValueType.Double, iStyleRef);
    internal static CellValue Create(int intValue, short iStyleRef)
        => new CellValue(intValue, CellValueType.Int, iStyleRef);
    internal static CellValue Create(long longValue, short iStyleRef)
        => new CellValue(longValue, CellValueType.Long, iStyleRef);

    internal static CellValue Create(ExcelErrorCode errorCodeValue)
        => new CellValue((int)errorCodeValue, CellValueType.Error, -1);

    private static CellValue Create(DBNull _, short iStyleRef)
        => new CellValue(CellValueType.IsDBNull, iStyleRef);

    internal static CellValue TryParseOrder(ReadOnlySpan<char> asSpan, CellStyle? style)
    {
        // Determine if the provided style corresponds to a date/time related formatting type.
        bool isDateStyle = style?.IsDateStyle ?? false;
        bool containsDecimal = asSpan.ContainsAny('.', 'E');
        short styleStyleXmlRef = style?.ExcelFormatId ?? -1;
        if (containsDecimal)
        {
            return PerformDecimalConversion(asSpan, styleStyleXmlRef, isDateStyle);
        }
        bool containsSign = asSpan[0] == '-';
        int asSpanLength = asSpan.Length;

        // ReSharper disable once ConvertIfStatementToSwitchStatement
        if (!containsSign && asSpanLength < 11)
        {
            int intVal = asSpan.IntParse();
            return isDateStyle
                ? CellValue.Create(DateTime.FromOADate(intVal), styleStyleXmlRef)
                : CellValue.Create(intVal, styleStyleXmlRef);
        }

        if (containsSign && asSpanLength == 12
            && int.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resultI)
           )
        {
            // -2,147,483,648 to 2,147,483,647 	Signed 32-bit integer
            return isDateStyle
                ? CellValue.Create(DateTime.FromOADate(resultI), styleStyleXmlRef)
                : CellValue.Create(resultI, styleStyleXmlRef);
        }

        if ((asSpanLength < 19 || (containsSign && asSpanLength == 20))
            && long.TryParse(asSpan, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out long resultL))
        {
            // -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 	Signed 64-bit integer
            return isDateStyle
                ? CellValue.Create(DateTime.FromOADate(resultL), styleStyleXmlRef)
                : CellValue.Create(resultL, styleStyleXmlRef);
        }

        return PerformDecimalConversion(asSpan, styleStyleXmlRef, isDateStyle);
    }

    private static CellValue PerformDecimalConversion(ReadOnlySpan<char> asSpan, short style, bool isDateStyle)
    {
        if (asSpan.TryDecimalParse(out decimal resultM))
        {
            // ±1.0 x 10-28 to ±7.9228 x 1028 	28-29 digits 	16 bytes
            return isDateStyle
                ? CellValue.Create(DateTime.FromOADate((double)resultM), style)
                : CellValue.Create(resultM, style);
        }
        if (asSpan.TryDoubleParse(out double resultD))
        {
            //   	±5.0 × 10−324 to ±1.7 × 10308 	~15-17 digits 	8 bytes
            return isDateStyle
                ? CellValue.Create(DateTime.FromOADate(resultD), style)
                : CellValue.Create(resultD, style);
        }
        // If the format is a date style, and parsing as decimal/double failed, attempt a direct double parse to OA date
        if (isDateStyle && double.TryParse(asSpan, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedD))
        {
            return CellValue.Create(DateTime.FromOADate(parsedD), style);
        }
        return CellValue.Create(new string(asSpan), style);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="CellValue"/> class.
    /// </summary>
    /// <param name="type">The type of the cell value.</param>
    /// <param name="iStyleRef">The style reference index.</param>
    private CellValue(CellValueType type, short iStyleRef)
    {
        _type = type;
        _iStyleRef = iStyleRef;
    }

    private CellValue(bool value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _b = value;
    }

    private CellValue(DateTime value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _dt = value;
    }

    private CellValue(string value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _s = value;
    }

    private CellValue(decimal value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _d = value;
    }

    private CellValue(double value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _db = value;
    }

    private CellValue(int value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _i = value;
    }

    private CellValue(long value, CellValueType type, short iStyleRef) : this(type, iStyleRef)
    {
        _l = value;
    }

    /// <summary>
    /// Gets the raw "Boxed" value of the cell.
    /// </summary>
    public object? BoxedValue => _type switch
    {
        CellValueType.Decimal => _d,
        CellValueType.Double => _db,
        CellValueType.Int => _i,
        CellValueType.Long => _l,
        CellValueType.String => _s,
        CellValueType.Bool => _b,
        CellValueType.DateTime => _dt,
        CellValueType.Error => (ExcelErrorCode)_i,
        CellValueType.IsDBNull => DBNull.Value,
        _ => null
    };

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
        if (_type == CellValueType.String)
        {
            return _s;
        }

        return ToString_Slow();
    }

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private string? ToString_Slow() =>
        _type switch
        {
            CellValueType.Bool => _b.ToString(s_invariantCultureCache),
            // Micro-optimization: Use cached CultureInfo instead of property access
            CellValueType.Decimal => _d.ToInvariantString(),
            CellValueType.Double => _db.ToInvariantString(),
            CellValueType.Int => _i.ToInvariantString(),
            CellValueType.Long => _l.ToInvariantString(),
            CellValueType.DateTime => _dt.ToString(s_invariantCultureCache),
            CellValueType.Error => ((ExcelErrorCode)_i).ToString(),
            CellValueType.IsDBNull => DBNull.Value.ToString(s_invariantCultureCache),
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
        // Try zero-allocation formatting using stackalloc buffer
        Span<char> buffer = stackalloc char[64]; // Adjust size as needed
        if (TryFormat(buffer, out int charsWritten, default, null))
        {
            builder.Append(buffer.Slice(0, charsWritten));
            return;
        }
        // Fallback to string-based formatting if buffer is too small
        builder.Append(ToString());
    }


    /// <summary>
    /// Gets the value of the cell as a <see cref="DateTime"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="DateTime"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public DateTime AsDateTime =>
        _type == CellValueType.DateTime
        ? _dt
        : AsDateTime_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private DateTime AsDateTime_Slow() =>
        _type switch
        {
            CellValueType.Decimal => DateTime.FromOADate((double)_d),
            CellValueType.Double => DateTime.FromOADate(_db),
            CellValueType.Long => DateTime.FromOADate(_l),
            CellValueType.Int => DateTime.FromOADate(_i),
            _ => double.TryParse(_s, out double val)
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
            ? _b
            : AsBoolean_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private bool AsBoolean_Slow() =>
        _type switch
        {
            CellValueType.Decimal => _d != 0,
            CellValueType.Double => _db != 0,
            CellValueType.Long => _l != 0,
            CellValueType.Int => _i != 0,
            CellValueType.Error => (ExcelErrorCode)AsInt32 != ExcelErrorCode.Null,
            CellValueType.IsDBNull => false,
            _ => int.TryParse(_s, out int val) ? val != 0 : _s!.BoolParse()
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
        _type == CellValueType.Int
            ? _i
            : AsInt32_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private int AsInt32_Slow() =>
        _type switch
        {
            CellValueType.Decimal => (int)_d,
            CellValueType.Double => (int)_db,
            CellValueType.Long => (int)_l,
            CellValueType.DateTime => (int)_dt.Ticks,
            CellValueType.Bool => _b ? 1 : 0,
            CellValueType.Error => _i,
            CellValueType.IsDBNull => 0,
            _ => int.Parse(_s!, NumberStyles.Integer, CultureInfo.InvariantCulture)
        };

    /// <summary>
    /// Gets the value of the cell as a <see cref="Int64"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Int64"/>.
    /// </exception>
    // Remove AggressiveOptimization from properties
    public long AsInt64 =>
        _type == CellValueType.Long
            ? _l
            : AsInt64_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private long AsInt64_Slow() =>
        _type switch
        {
            CellValueType.Decimal => (long)_d,
            CellValueType.Double => (long)_db,
            CellValueType.Int => (long)_i,
            CellValueType.DateTime => (long)_dt.Ticks,
            CellValueType.Bool => _b ? 1 : 0,
            CellValueType.Error => (long)_i,
            CellValueType.IsDBNull => 0L,
            _ => long.Parse(_s!, NumberStyles.Integer, CultureInfo.InvariantCulture)
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
        _type == CellValueType.Double
            ? _db
            : AsDouble_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private double AsDouble_Slow() =>
        _type switch
        {
            CellValueType.Decimal => (double)_d,
            CellValueType.Long => (double)_l,
            CellValueType.Int => (double)_i,
            CellValueType.DateTime => (double)_dt.Ticks,
            CellValueType.Bool => _b ? 1 : 0,
            CellValueType.Error => (double)_i,
            CellValueType.IsDBNull => 0.0,
            _ => double.Parse(_s!, NumberStyles.Float, CultureInfo.InvariantCulture)
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
        _type == CellValueType.Decimal
            ? _d
            : AsDecimal_Slow();

    [MethodImpl(MethodImplOptions.NoInlining)] // Keep hot path small
    private decimal AsDecimal_Slow() =>
        _type switch
        {
            CellValueType.Double => (decimal)_db,
            CellValueType.Long => (decimal)_l,
            CellValueType.Int => (decimal)_i,
            CellValueType.DateTime => (decimal)_dt.Ticks,
            CellValueType.Bool => _b ? 1m : 0m,
            CellValueType.Error => (decimal)_i,
            CellValueType.IsDBNull => 0m,
            _ => decimal.Parse(_s!, NumberStyles.Currency, CultureInfo.InvariantCulture)
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
            CellValueType.Decimal => FormatNumericWithNumberFormat(_d, formatCode, type),
            CellValueType.Double => FormatNumericWithNumberFormat(_db, formatCode, type),
            CellValueType.Long => FormatNumericWithNumberFormat((decimal)_l, formatCode, type),
            CellValueType.Int => FormatNumericWithNumberFormat((decimal)_i, formatCode, type),
            CellValueType.DateTime => FormatDateTimeWithNumberFormat(AsDateTime, formatCode, type),
            CellValueType.Bool => _d != 0 ? bool.TrueString : bool.FalseString, // What if the style is to upper case ?
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
    private static string FormatNumericWithNumberFormat(decimal value, string formatCode, FormattingType type)
    {
        // Ensure we're handling FormattingType.Number
        if (type != FormattingType.Number)
        {
            return value.ToString(s_invariantCultureCache);
        }

        return formatCode switch
        {
            // General and text formats
            "General" => value.ToString(s_invariantCultureCache),
            "@" => value.ToString(s_invariantCultureCache),

            // Basic integer formats
            "0" => Math.Round(value).ToString(s_invariantCultureCache),

            // Decimal formats
            "0.00" => value.ToString("F2", s_invariantCultureCache),
            "0.0" => value.ToString("F1", s_invariantCultureCache),

            // Thousand separator formats
            "#,##0" => Math.Round(value).ToString("N0", s_invariantCultureCache),
            "#,##0.0" => value.ToString("N1", s_invariantCultureCache),
            "#,##0.00" => value.ToString("N2", s_invariantCultureCache),

            // With brackets (negative in parentheses)
            "#,##0;(#,##0)" => FormatNumberWithNegativeParentheses(value, "N0"),
            "#,##0.0;(#,##0.0)" => FormatNumberWithNegativeParentheses(value, "N1"),
            "#,##0.00;(#,##0.00)" => FormatNumberWithNegativeParentheses(value, "N2"),

            // With red color indicator for negatives
            "#,##0;[Red](#,##0)" => FormatNumberWithNegativeParentheses(value, "N0"),
            "#,##0.00;[Red](#,##0.00)" => FormatNumberWithNegativeParentheses(value, "N2"),

            // Percentage formats
            "0%" => FormatPercentage(value, "F0"),
            "0.0%" => FormatPercentage(value, "F1"),
            "0.00%" => FormatPercentage(value, "F2"),

            // Fallback for complex formats
            _ => FormatNumericWithNumberFormat((double)value, formatCode, type)
        };
    }

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
            return value.ToString(s_invariantCultureCache);
        }

        return formatCode switch
        {
            // General and text formats
            "General" => value.ToString(s_invariantCultureCache),
            "@" => value.ToString(s_invariantCultureCache),

            // Basic integer formats
            "0" => Math.Round(value).ToString(s_invariantCultureCache),

            // Decimal formats
            "0.00" => value.ToString("F2", s_invariantCultureCache),
            "0.0" => value.ToString("F1", s_invariantCultureCache),

            // Thousand separator formats
            "#,##0" => Math.Round(value).ToString("N0", s_invariantCultureCache),
            "#,##0.0" => value.ToString("N1", s_invariantCultureCache),
            "#,##0.00" => value.ToString("N2", s_invariantCultureCache),

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
            "0%" => FormatPercentage(value, "F0"),
            "0.0%" => FormatPercentage(value, "F1"),
            "0.00%" => FormatPercentage(value, "F2"),

            // Scientific notation
            "0.00E+00" => value.ToString("E2", s_invariantCultureCache),
            "0.00E+0" => value.ToString("E2", s_invariantCultureCache),
            "0.00E0" => value.ToString("E2", s_invariantCultureCache),
            "##0.0E0" => value.ToString("E1", s_invariantCultureCache),
            "##0.0E+0" => value.ToString("E1", s_invariantCultureCache),
            "##0.0E+00" => value.ToString("E1", s_invariantCultureCache),

            // Fraction formats
            "# ?/?" => FormatFraction(value, 1),
            "# ??/??" => FormatFraction(value, 2),

            // CJK formats (treated as numbers)
            "[DBNum1][$-804]0" => value.ToString("F0", s_invariantCultureCache),
            "[DBNum1][$-804]0.00" => value.ToString("F2", s_invariantCultureCache),
            "[DBNum4][$-804]0" => value.ToString("F0", s_invariantCultureCache),

            // Default: return as double with invariant culture
            _ => FormatCustomNumber(value, formatCode)
        };
    }

    /// <summary>
    /// Formats a number with negative values in parentheses.
    /// </summary>
    private static string FormatNumberWithNegativeParentheses(decimal value, string format)
    {
        if (value < 0)
        {
            Span<char> buffer = stackalloc char[128];
            buffer[0] = '(';
            if (Math.Abs(value).TryFormat(buffer.Slice(1), out int written, format, s_invariantCultureCache))
            {
                buffer[written + 1] = ')';
                return new string(buffer.Slice(0, written + 2));
            }
        }
        return value.ToString(format, s_invariantCultureCache);
    }

    /// <summary>
    /// Formats a number with negative values in parentheses.
    /// </summary>
    private static string FormatNumberWithNegativeParentheses(double value, string format)
    {
        if (value < 0)
        {
            Span<char> buffer = stackalloc char[128];
            buffer[0] = '(';
            if (Math.Abs(value).TryFormat(buffer.Slice(1), out int written, format, s_invariantCultureCache))
            {
                buffer[written + 1] = ')';
                return new string(buffer.Slice(0, written + 2));
            }
        }
        return value.ToString(format, s_invariantCultureCache);
    }

    /// <summary>
    /// Formats a number in accounting style with alignment spacing.
    /// </summary>
    private static string FormatAccountingNumber(double value, int decimals)
    {
        Span<char> format = stackalloc char[12];
        format[0] = 'N';
        if (!decimals.TryFormat(format.Slice(1), out int fWritten, default, s_invariantCultureCache))
        {
            // Fallback
            return value.ToString("N" + decimals, s_invariantCultureCache);
        }

        ReadOnlySpan<char> formatSpan = format.Slice(0, 1 + fWritten);
        Span<char> buffer = stackalloc char[128];
        int pos = 0;

        if (value < 0)
        {
            buffer[pos++] = '(';
            if (Math.Abs(value).TryFormat(buffer.Slice(pos), out int vWritten, formatSpan, s_invariantCultureCache))
            {
                pos += vWritten;
                buffer[pos++] = ')';
                return new string(buffer.Slice(0, pos));
            }
        }
        else
        {
            buffer[pos++] = ' ';
            if (value.TryFormat(buffer.Slice(pos), out int vWritten, formatSpan, s_invariantCultureCache))
            {
                pos += vWritten;
                buffer[pos++] = ' ';
                return new string(buffer.Slice(0, pos));
            }
        }

        return value.ToString(formatSpan.ToString(), s_invariantCultureCache);
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
            return Math.Round(value).ToString(s_invariantCultureCache);
        }

        // Stack allocation for fraction formatting
        Span<char> buffer = stackalloc char[32];
        DefaultInterpolatedStringHandler handler = new(1, 2, s_invariantCultureCache, buffer);

        if (intPart != 0)
        {
            handler.AppendFormatted((long)Math.Truncate(intPart));
            handler.AppendLiteral(" ");
        }

        handler.AppendFormatted(numerator);
        handler.AppendLiteral("/");
        handler.AppendFormatted(denominator);

        return handler.ToStringAndClear();
    }

    /// <summary>
    /// Formats a numeric value as a percentage with the specified decimal places.
    /// </summary>
    private static string FormatPercentage(decimal value, string formatSpec) => FormatPercentage((double)value, formatSpec);

    /// <summary>
    /// Formats a numeric value as a percentage with the specified decimal places.
    /// </summary>
    private static string FormatPercentage(double value, string formatSpec)
    {
        Span<char> buffer = stackalloc char[64];
        double percentValue = value * 100;

        if (percentValue.TryFormat(buffer, out int written, formatSpec, s_invariantCultureCache))
        {
            DefaultInterpolatedStringHandler handler = new(1, 1, s_invariantCultureCache, buffer);
            handler.AppendLiteral(new string(buffer.Slice(0, written)));
            handler.AppendLiteral("%");
            return handler.ToStringAndClear();
        }

        return percentValue.ToString(formatSpec, s_invariantCultureCache) + "%";
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
        ReadOnlySpan<char> span = formatCode.AsSpan();
        // Check for percentage format in the code
        if (span.Contains('%'))
        {
            // Count decimal places in the format
            int decimalPlaces = 0;
            int dotIndex = span.IndexOf('.');
            if (dotIndex >= 0)
            {
                ReadOnlySpan<char> afterDot = span.Slice(dotIndex + 1);
                foreach (char c in afterDot)
                {
                    if (c == '0')
                    {
                        decimalPlaces++;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            Span<char> formatBuffer = stackalloc char[12];
            formatBuffer[0] = 'F';
            if (decimalPlaces.TryFormat(formatBuffer.Slice(1), out int fWritten, default, s_invariantCultureCache))
            {
                Span<char> resultBuffer = stackalloc char[128];
                if ((value * 100).TryFormat(resultBuffer, out int vWritten, formatBuffer.Slice(0, 1 + fWritten), s_invariantCultureCache))
                {
                    resultBuffer[vWritten] = '%';
                    return new string(resultBuffer.Slice(0, vWritten + 1));
                }
            }
            return (value * 100).ToString("F" + decimalPlaces, s_invariantCultureCache) + "%";
        }

        // Check for scientific notation
        if (span.ContainsAny("Ee"))
        {
            int eIndex = Math.Max(span.IndexOf('E'), span.IndexOf('e'));
            int decimalPlaces = 2; // default
            if (eIndex > 0)
            {
                ReadOnlySpan<char> beforeE = span.Slice(0, eIndex);
                int dotIndex = beforeE.LastIndexOf('.');
                if (dotIndex >= 0)
                {
                    decimalPlaces = beforeE.Length - dotIndex - 1;
                }
            }

            Span<char> formatBuffer = stackalloc char[12];
            formatBuffer[0] = 'E';
            if (decimalPlaces.TryFormat(formatBuffer.Slice(1), out int fWritten, default, s_invariantCultureCache))
            {
                Span<char> resultBuffer = stackalloc char[128];
                if (value.TryFormat(resultBuffer, out int vWritten, formatBuffer.Slice(0, 1 + fWritten), s_invariantCultureCache))
                {
                    return new string(resultBuffer.Slice(0, vWritten));
                }
            }
            return value.ToString("E" + decimalPlaces, s_invariantCultureCache);
        }

        // Count decimal places from format code
        int dotPos = span.IndexOf('.');
        int decimals = 0;
        if (dotPos >= 0)
        {
            ReadOnlySpan<char> afterDot = span.Slice(dotPos + 1);
            foreach (char c in afterDot)
            {
                if (c is '0' or '#')
                {
                    decimals++;
                }
            }
        }

        // Check if thousands separator is present
        if (span.Contains(','))
        {
            Span<char> formatBuffer = stackalloc char[12];
            formatBuffer[0] = 'N';
            if (decimals.TryFormat(formatBuffer.Slice(1), out int fWritten, default, s_invariantCultureCache))
            {
                Span<char> resultBuffer = stackalloc char[128];
                if (value.TryFormat(resultBuffer, out int vWritten, formatBuffer.Slice(0, 1 + fWritten), s_invariantCultureCache))
                {
                    return new string(resultBuffer.Slice(0, vWritten));
                }
            }
            return value.ToString("N" + decimals, s_invariantCultureCache);
        }

        // Default fixed-point format
        if (decimals > 0)
        {
            Span<char> formatBuffer = stackalloc char[12];
            formatBuffer[0] = 'F';
            if (decimals.TryFormat(formatBuffer.Slice(1), out int fWritten, default, s_invariantCultureCache))
            {
                Span<char> resultBuffer = stackalloc char[128];
                if (value.TryFormat(resultBuffer, out int vWritten, formatBuffer.Slice(0, 1 + fWritten), s_invariantCultureCache))
                {
                    return new string(resultBuffer.Slice(0, vWritten));
                }
            }
            return value.ToString("F" + decimals, s_invariantCultureCache);
        }

        return Math.Round(value).ToString(s_invariantCultureCache);
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

        Span<char> buffer = stackalloc char[32];
        int pos = 0;
        if (totalHours.TryFormat(buffer, out int written, default, s_invariantCultureCache))
        {
            pos += written;
            buffer[pos++] = ':';
            if (minutes.TryFormat(buffer.Slice(pos), out written, "D2", s_invariantCultureCache))
            {
                pos += written;
                buffer[pos++] = ':';
                if (seconds.TryFormat(buffer.Slice(pos), out written, "D2", s_invariantCultureCache))
                {
                    pos += written;
                    return new string(buffer.Slice(0, pos));
                }
            }
        }

        // Fallback
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
                    // If preceded by a digit or date separator
                    char charBefore = result[i - 1];
                    bool isProbablyMonth = i > 0
                                           && (char.IsDigit(charBefore)
                                               || charBefore == '/'
                                               || charBefore == '-'
                                               || charBefore == '.'
                                           );

                    // If followed by digit, date separator, or date character
                    if (i + 2 < result.Length)
                    {
                        char next = result[i + 2];
                        if (char.IsDigit(next)
                            || next == '/' || next == '-' || next == '.' || next == 'd' || next == 'y' || next == 'D' || next == 'Y')
                        {
                            isProbablyMonth = true;
                        }
                    }

                    sb.Append(isProbablyMonth ? "MM" : "mm");
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
    public override int GetHashCode() =>
        _type switch
        {
            CellValueType.IsDBNull => (int)CellValueType.IsDBNull,
            CellValueType.String => HashCode.Combine(_type, _s, _iStyleRef),
            _ => HashCode.Combine(_type, AsDecimal, _iStyleRef)
        };

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
    public bool Equals(CellValue? other)
    {
        if (other is null || _type != other._type)
        {
            return false;
        }
        return _type switch
        {
            CellValueType.Bool => _b == other._b,
            CellValueType.Decimal => _d == other._d,
            CellValueType.Double => _db == other._db,
            CellValueType.Long => _l == other._l,
            CellValueType.Int => _i == other._i,
            CellValueType.DateTime => _dt == other._dt,
            CellValueType.IsDBNull => true,
            CellValueType.String => _s == other._s,
            _ => false
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
    public static bool operator ==(CellValue? left, CellValue? right)
        => left?.Equals(right) ?? right is null;

    /// <summary>
    /// Determines whether two<see cref = "CellValue" /> instances are not equal.
    /// </summary>
    /// <param name = "left" > The first<see cref = "CellValue" /> to compare.</param>
    /// <param name = "right" > The second<see cref = "CellValue" /> to compare.</param>
    /// <returns>
    /// <c>true</c> if the specified<see cref="CellValue"/> instances are not equal; otherwise, <c>false</c>.
    /// </returns>
    public static bool operator !=(CellValue? left, CellValue? right)
        => !(left?.Equals(right) ?? right is null);

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
        if (_type == CellValueType.String)
        {
            return TryFormatString(destination, out charsWritten);
        }

        // Dispatch to type-specific handling
        return TryFormat_Slow(destination, out charsWritten);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private bool TryFormatString(Span<char> destination, out int charsWritten)
    {
        if (_s == null)
        {
            charsWritten = 0;
            return true;
        }

        if (destination.Length < _s.Length)
        {
            charsWritten = 0;
            return false;
        }

        _s.AsSpan().CopyTo(destination);
        charsWritten = _s.Length;
        return true;
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private bool TryFormat_Slow(Span<char> destination, out int charsWritten)
    {
        switch (_type)
        {
            case CellValueType.Decimal:
                return _d.TryFormat(destination, out charsWritten, default, s_invariantCultureCache);

            case CellValueType.Double:
                return _db.TryFormat(destination, out charsWritten, default, s_invariantCultureCache);

            case CellValueType.Long:
                return _l.TryFormat(destination, out charsWritten, default, s_invariantCultureCache);

            case CellValueType.Int:
                return _i.TryFormat(destination, out charsWritten, default, s_invariantCultureCache);

            case CellValueType.Bool:
                {
                    string boolStr = _b ? bool.TrueString : bool.FalseString;
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
                return AsDateTime.TryFormat(destination, out charsWritten, default, s_invariantCultureCache);

            case CellValueType.Error:
                {
                    string errorStr = ((ExcelErrorCode)AsInt32).ToString();
                    if (destination.Length < errorStr.Length)
                    {
                        charsWritten = 0;
                        return false;
                    }
                    errorStr.AsSpan().CopyTo(destination);
                    charsWritten = errorStr.Length;
                    return true;
                }

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

    #region SIMD-Accelerated Filtering Operations

    /// <summary>
    /// Filters a batch of cell values by type using zero-allocation operations.
    /// </summary>
    /// <param name="cells">Input array of cell values to filter.</param>
    /// <param name="output">Pre-allocated output array for matching cells.</param>
    /// <param name="targetType">The cell value type to filter by.</param>
    /// <returns>Count of cells written to output array.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static int FilterByCellType(ReadOnlySpan<CellValue> cells, Span<CellValue> output, CellValueType targetType)
    {
        int outputIndex = 0;
        foreach (CellValue cell in cells)
        {
            if (outputIndex >= output.Length)
            {
                break;
            }

            if (cell._type == targetType)
            {
                output[outputIndex++] = cell;
            }
        }
        return outputIndex;
    }

    /// <summary>
    /// Checks if this cell value is numeric (optimized for filtering operations).
    /// </summary>
    public bool IsDecimal
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.Decimal;
    }
    /// <summary>
    /// Checks if this cell value is numeric (optimized for filtering operations).
    /// </summary>
    public bool IsDouble
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.Double;
    }
    /// <summary>
    /// Checks if this cell value is numeric (optimized for filtering operations).
    /// </summary>
    public bool IsLong
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.Long;
    }
    /// <summary>
    /// Checks if this cell value is numeric (optimized for filtering operations).
    /// </summary>
    public bool IsInt
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.Int;
    }

    /// <summary>
    /// Checks if this cell value is a string (optimized for filtering operations).
    /// </summary>
    public bool IsString
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.String;
    }

    /// <summary>
    /// Checks if this cell value is a boolean (optimized for filtering operations).
    /// </summary>
    public bool IsBoolean
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.Bool;
    }

    /// <summary>
    /// Checks if this cell value is a datetime (optimized for filtering operations).
    /// </summary>
    public bool IsDateTime
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.DateTime;
    }

    /// <summary>
    /// Checks if this cell value is an error (optimized for filtering operations).
    /// </summary>
    public bool IsError
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.Error;
    }

    /// <summary>
    /// Checks if this cell value is DBNull (optimized for filtering operations).
    /// </summary>
    public bool IsDBNull
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => _type == CellValueType.IsDBNull;
    }
    #endregion

}

