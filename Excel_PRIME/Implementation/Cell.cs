using System;
using System.Buffers;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{RawValue.ToString(),raw}")]
internal class Cell : ICell
{
    private const int BufferSize = 64;

    public static async Task<Cell> ConstructCellAsync(XmlReader reader, InstanceContext instanceContext, ReaderAtoms readerAtoms)
    {
        string address = string.Empty;
        CellType type = CellType.Unknown;
        object? value = null;
        char[] buffer = ArrayPool<char>.Shared.Rent(BufferSize);
        try
        {
            int len;

            void ReadValue()
            {
                len = reader.ReadValueChunk(buffer, 0, BufferSize);
            }

            while (reader.MoveToNextAttribute())
            {
                // Retrieve the atomized name directly.
                string currentAttributeName = reader.LocalName;
                if (Object.ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
                {
                    address = reader.Value;
                }
                else if (Object.ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
                {
                    ReadValue();
                    type = GetCellType(buffer, len);
                }
                else if (Object.ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom))
                {
                    // TODO: the style, therefore converting into time only etc.
                    //ReadValue();
                    //style = GetStyleOffset(buffer, len);
                }
            }

            if (await reader.ReadAsync().ConfigureAwait(false)
                && !reader.IsEmptyElement
                && Object.ReferenceEquals(reader.LocalName, readerAtoms.vRefAtom)
               )
            {
                // Move to data
                await reader.ReadAsync().ConfigureAwait(false);

                switch (type)
                {
                    case CellType.Unknown:
                        value = reader.ReadString();
                        break;
                    case CellType.Numeric:
                        ReadValue();
                        value = TryParseOrder(instanceContext.Options.CellConversionType, buffer.AsSpan(0, len));
                        break;
                    case CellType.String:
                        value = reader.ReadString();
                        break;
                    case CellType.SharedString:
                        ReadValue();
                        value = instanceContext.SharedStrings[buffer.AsSpan(0, len).IntParse()];
                        break;
                    case CellType.InlineString:
                        value = reader.ReadString();
                        break;
                    case CellType.Boolean:
                        ReadValue();
                        value = buffer[0] == '0';
                        break;
                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = reader.ReadString();
                        break;
                    case CellType.Date:
                        ReadValue();
                        if (double.TryParse(buffer.AsSpan(0, len), NumberStyles.Number,
                                CultureInfo.InvariantCulture, out double dateTimeValue))
                        {
                            value = DateTime.FromOADate(dateTimeValue);
                        }
                        else if (DateTime.TryParse(buffer, out DateTime result))
                        {
                            value = result;
                        }
                        else
                        {
                            value = new string(buffer.AsSpan(0, len));
                        }

                        break;
                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer);
        }
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        (int _, int col, ReadOnlyMemory<char> colName) = address.GetRowColNumbers();
        return new Cell
        {
            ColumnLetters = colName,
            //RowNumber = row;
            ExcelColumnOffset = col,
            RawExcelType = type,
            RawValue = value
        };
    }

    //[MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static object? TryParseOrder(CellConversion optionsCellConversionType, ReadOnlySpan<char> asSpan)
    {
        if (asSpan.Length == 0)
        {
            return null;
        }

        return optionsCellConversionType switch
        {
            CellConversion.None => new ReadOnlyMemory<char>(asSpan.ToArray()),
            CellConversion.Number => PerformNumberLiteralConversion(asSpan),
            CellConversion.NumberAndDates => // TODO
                PerformNumberLiteralConversion(asSpan),
            CellConversion.FromStyles => // TODO
                PerformNumberLiteralConversion(asSpan),
            _ => new ReadOnlyMemory<char>(asSpan.ToArray())
        };
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static object PerformNumberLiteralConversion(ReadOnlySpan<char> asSpan)
    {
        bool containsDecimal = asSpan.Contains('.');
        if (containsDecimal)
        {
            //float
            // ±1.5 x 10−45 to ±3.4 x 1038 	~6-9 digits 	4 bytes
            return PerformSimpleConversion(asSpan);
        }

        // ReSharper disable once ConvertIfStatementToSwitchStatement
        if (asSpan.Length < 12
            && int.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resultI)
           )
        {
            // -2,147,483,648 to 2,147,483,647 	Signed 32-bit integer
            return resultI;
        }

        if (asSpan.Length < 20
            && long.TryParse(asSpan, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out long resultL))
        {
            // -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 	Signed 64-bit integer
            return resultL;
        }

        if (asSpan.Length > 18
            && BigInteger.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture,
                out BigInteger resultBI))
        {
            return resultBI;
        }

        return PerformSimpleConversion(asSpan);
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static object PerformSimpleConversion(ReadOnlySpan<char> asSpan)
    {
        if (decimal.TryParse(asSpan, NumberStyles.Currency,
                CultureInfo.InvariantCulture, out decimal resultM))
        {   // ±1.0 x 10-28 to ±7.9228 x 1028 	28-29 digits 	16 bytes
            return resultM;
        }
        if (double.TryParse(asSpan, NumberStyles.Float,
                CultureInfo.InvariantCulture, out double resultD))
        {   //  	±5.0 × 10−324 to ±1.7 × 10308 	~15-17 digits 	8 bytes
            return resultD;
        }
        return new string(asSpan);
    }


    /// <InheritDoc />
    public object? RawValue { get; private init; }

    /// <InheritDoc />
    public CellType RawExcelType { get; private init; }

    /// <InheritDoc />
    public ReadOnlyMemory<char> ColumnLetters { get; private init; }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; private init; }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static CellType GetCellType(char[] b, int l)
    {
        if (l == 0)
        {
            // Default type is Numeric and some Excel do not write this
            return CellType.Numeric;
        }

        return b[0] switch
        {
            'b' => CellType.Boolean,
            'e' => CellType.Error,
            's' => l == 1 ? CellType.SharedString : CellType.String,
            'i' => CellType.InlineString,
            'd' => CellType.Date,
            'n' => CellType.Numeric,
            _ => throw new InvalidDataException()
        };
    }
}