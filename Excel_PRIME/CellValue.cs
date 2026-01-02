using System;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME;

/// <summary>
/// Represents a strongly-typed cell value with custom ToString conversion.
/// </summary>
public struct CellValue : IEquatable<CellValue>
{
    private enum CellValueType
    {
        Unknown,
        Numeric,
        String,
        Bool,
        Error,
        DateTime
    }

    private string? _strValue; // Has to be on it's own due to reference type
    private readonly CellValueType _type;
    private readonly BclValue _value;

    [StructLayout(LayoutKind.Explicit)]
    private struct BclValue
    {
        [FieldOffset(0)] public bool _boolValue;
        [FieldOffset(0)] public double _doubleValue;
        [FieldOffset(0)] public DateTime _dateTimeValue;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal CellValue(string? strValue)
    {
        _strValue = strValue;
        _type = CellValueType.String;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal CellValue(bool boolValue)
    {
        _value = new BclValue { _boolValue = boolValue };
        _type = CellValueType.Bool;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal CellValue(double doubleValue)
    {
        _value = new BclValue { _doubleValue = doubleValue };
        _type = CellValueType.Numeric;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal CellValue(DateTime dateTimeValue)
    {
        _value = new BclValue { _dateTimeValue = dateTimeValue };
        _type = CellValueType.DateTime;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal CellValue(ExcelErrorCode errorCodeValue)
    {
        _value = new BclValue { _doubleValue = (int)errorCodeValue };
        _type = CellValueType.Error;
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
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public override string? ToString()
    {
        _strValue ??= _type switch
        {
            CellValueType.Unknown => null,
            CellValueType.Bool => _value._boolValue ? bool.TrueString : bool.FalseString,
            CellValueType.Numeric => _value._doubleValue.ToString(CultureInfo.InvariantCulture),
            CellValueType.DateTime => _value._dateTimeValue.ToString(CultureInfo.InvariantCulture),
            CellValueType.Error => ((ExcelErrorCode)_value._doubleValue).ToString()
        };

        return _strValue;
    }

    /// <summary>
    /// Gets the raw "Boxed" value of the cell.
    /// </summary>
    public object? BoxedValue =>
        _type switch
        {
            CellValueType.Unknown => null,
            CellValueType.Bool => _value._boolValue,
            CellValueType.Numeric => _value._doubleValue,
            CellValueType.DateTime => _value._dateTimeValue,
            CellValueType.Error => (ExcelErrorCode)_value._doubleValue,
            CellValueType.String => _strValue,
            _ => null
        };


    /// <summary>
    /// Gets the value of the cell as a <see cref="DateTime"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="DateTime"/>.
    /// </exception>
    public DateTime AsDateTime
    {
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
        get
        {
            switch (_type)
            {
                case CellValueType.DateTime:
                    return _value._dateTimeValue;
                case CellValueType.Numeric:
                    return DateTime.FromOADate(_value._doubleValue);
                default:
                    return double.TryParse(_strValue, out double val)
                        ?// Excel stores the DateTime as a double OADate
                         DateTime.FromOADate(val)
                        : DateTime.Parse(ToString()!, CultureInfo.InvariantCulture);
            }
        }
    }

    /// <summary>
    /// Gets the value of the cell as a <see cref="bool"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="bool"/>.
    /// </exception>
    public bool AsBoolean
    {
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
        get
        {
            switch (_type)
            {
                case CellValueType.DateTime:
                    return _value._dateTimeValue.Ticks != 0;
                case CellValueType.Bool:
                    return _value._boolValue;
                case CellValueType.Error:
                    return (ExcelErrorCode)_value._doubleValue != ExcelErrorCode.Null;
                case CellValueType.Numeric:
                    return _value._doubleValue != 0;
                default:
                    return int.TryParse(_strValue, out int val)
                        ?// Sometimes Excel stores the Boolean as an int `1` or `0`
                        val != 0
                        : Convert.ToBoolean(_strValue!);
            }
        }
    }
    
    /// <summary>
    /// Gets the value of the cell as a <see cref="Int32"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Int32"/>.
    /// </exception>
    public int AsInt32
    {
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
        get
        {
            switch (_type)
            {
                case CellValueType.DateTime:
                    return (int)_value._dateTimeValue.Ticks;
                case CellValueType.Bool:
                    return _value._boolValue?1:0;
                case CellValueType.Error:
                case CellValueType.Numeric:
                    return (int)_value._doubleValue;
                default:
                    {
                        ReadOnlySpan<char> asSpan = _strValue!.AsSpan();
                        return asSpan[0] != '-'
                            ? asSpan.IntParse()
                            : int.Parse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture);
                    };
            }
        }
    }

    /// <summary>
    /// Gets the value of the cell as a <see cref="Int64"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Int64"/>.
    /// </exception>
    public long AsInt64
    {
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
        get
        {
            switch (_type)
            {
                case CellValueType.DateTime:
                    return _value._dateTimeValue.Ticks;
                case CellValueType.Bool:
                    return _value._boolValue ? 1 : 0;
                case CellValueType.Error:
                case CellValueType.Numeric:
                    return (long)_value._doubleValue;
                default:
                    return long.Parse(_strValue, NumberStyles.Integer, CultureInfo.InvariantCulture);
            }
        }
    }

    /// <summary>
    /// Gets the value of the cell as a <see cref="double"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="double"/>.
    /// </exception>
    public double AsDouble
    {
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
        get
        {
            switch (_type)
            {
                case CellValueType.DateTime:
                    return _value._dateTimeValue.ToOADate();
                case CellValueType.Bool:
                    return _value._boolValue ? 1 : 0;
                case CellValueType.Error:
                case CellValueType.Numeric:
                    return _value._doubleValue;
                default:
                    return double.Parse(_strValue, NumberStyles.Float, CultureInfo.InvariantCulture);
            }
        }
    }

    /// <summary>
    /// Gets the value of the cell as a <see cref="Decimal"/> object, if possible.
    /// </summary>
    /// <exception cref="FormatException">
    /// Thrown if the cell value cannot be parsed as a valid <see cref="Decimal"/>.
    /// </exception>
    public decimal AsDecimal
    {
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
        get
        {
            switch (_type)
            {
                case CellValueType.DateTime:
                    return (decimal)_value._dateTimeValue.ToOADate();
                case CellValueType.Bool:
                    return _value._boolValue ? 1 : 0;
                case CellValueType.Error:
                case CellValueType.Numeric:
                    return (decimal)_value._doubleValue;
                default:
                    return decimal.Parse(_strValue, NumberStyles.Currency, CultureInfo.InvariantCulture);
            }
        }
    }
    
    /// <summary>
     /// Determines whether the specified object is equal to the current <see cref="CellValue"/> instance.
     /// </summary>
     /// <param name="obj">The object to compare with the current <see cref="CellValue"/> instance.</param>
     /// <returns>
     /// <c>true</c> if the specified object is a <see cref="CellValue"/> and has the same type and value as the current instance; otherwise, <c>false</c>.
     /// </returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public override bool Equals(object? obj) => obj is CellValue other && Equals(other);

    /// <summary>
    /// Returns the hash code for the current <see cref="CellValue"/> instance.
    /// </summary>
    /// <returns>
    /// A 32-bit signed integer hash code that represents the current <see cref="CellValue"/>.
    /// </returns>
    /// <remarks>
    /// The hash code is computed based on the type of the cell value and its associated data,
    /// ensuring that equal <see cref="CellValue"/> instances produce the same hash code.
    /// </remarks>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public override int GetHashCode() => HashCode.Combine(_type, _value._boolValue, _value._doubleValue, _value._dateTimeValue, _strValue);

    /// <summary>
    /// Determines whether the current <see cref="CellValue"/> instance is equal to another <see cref="CellValue"/> instance.
    /// </summary>
    /// <param name="other">The <see cref="CellValue"/> instance to compare with the current instance.</param>
    /// <returns>
    /// <c>true</c> if the current instance and the <paramref name="other"/> instance are equal; otherwise, <c>false</c>.
    /// </returns>
    /// <remarks>
    /// Equality is determined based on the <see cref="CellType"/> and the value associated with it.
    /// For example, numeric values are compared numerically, strings are compared using string equality,
    /// and dates are compared using date-time equality.
    /// </remarks>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public bool Equals(CellValue other)
    {
        if (_type != other._type)
        {
            return false;
        }

        return _type switch
        {
            CellValueType.Bool => _value._boolValue == other._value._boolValue,
            CellValueType.Numeric => _value._doubleValue == other._value._doubleValue,
            CellValueType.DateTime => _value._boolValue == other._value._boolValue,
            //CellValueType.String => _strValue == other._strValue,
            _ => _strValue == other._strValue
        };
    }

    /// <summary>
    /// Determines whether two <see cref="CellValue"/> instances are equal.
    /// </summary>
    /// <param name="left">The first <see cref="CellValue"/> to compare.</param>
    /// <param name="right">The second <see cref="CellValue"/> to compare.</param>
    /// <returns>
    /// <c>true</c> if the specified <see cref="CellValue"/> instances are equal; otherwise, <c>false</c>.
    /// </returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static bool operator ==(CellValue left, CellValue right) => left.Equals(right);

    /// <summary>
    /// Determines whether two <see cref="CellValue"/> instances are not equal.
    /// </summary>
    /// <param name="left">The first <see cref="CellValue"/> to compare.</param>
    /// <param name="right">The second <see cref="CellValue"/> to compare.</param>
    /// <returns>
    /// <c>true</c> if the specified <see cref="CellValue"/> instances are not equal; otherwise, <c>false</c>.
    /// </returns>
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static bool operator !=(CellValue left, CellValue right) => !(left == right);
}