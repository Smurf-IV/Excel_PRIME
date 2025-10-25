using System;
using System.Buffers;
using System.Globalization;
using System.IO;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class Cell : ICell
{
    private const int BufferSize = 64;

    public static async Task<Cell> ConstructCellAsync(XmlReader reader, ISharedString sharedStrings)
    {
        string rRef = reader.NameTable.Add("r");
        string tRef = reader.NameTable.Add("t");
        string vRef = reader.NameTable.Add("v");
        string sRef = reader.NameTable.Add("s");
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
                if (Object.ReferenceEquals(currentAttributeName, rRef))
                {
                    address = reader.Value;
                }
                else if (Object.ReferenceEquals(currentAttributeName, tRef))
                {
                    ReadValue();
                    type = GetCellType(buffer, len);
                }
                else if (Object.ReferenceEquals(currentAttributeName, sRef))
                {
                    // TODO: the style, therefore converting into time only etc.
                    //ReadValue();
                    //style = GetStyleOffset(buffer, len);
                }
            }

            if (await reader.ReadAsync().ConfigureAwait(false)
                && !reader.IsEmptyElement
                && Object.ReferenceEquals(reader.LocalName, vRef)
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
                        value = TryParseOrder(buffer.AsSpan(0, len));
                        break;
                    case CellType.String:
                        value = reader.ReadString();
                        break;
                    case CellType.SharedString:
                        ReadValue();
                        value = sharedStrings[buffer.AsSpan(0, len).IntParse()];
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

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static object? TryParseOrder(ReadOnlySpan<char> asSpan)
    {
        if (asSpan.Length == 0)
        {
            return null;
        }

        bool containsDecimal = asSpan.Contains('.');
        if (!containsDecimal
            && asSpan.Length < 11
            && int.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resultI)
            )
        {   // -2,147,483,648 to 2,147,483,647 	Signed 32-bit integer
            return resultI;
        }
        if (!containsDecimal
            && asSpan.Length > 9
            && long.TryParse(asSpan, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out long resultL))
        {   // -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 	Signed 64-bit integer
            return resultL;
        }
        if (!containsDecimal
            && BigInteger.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out BigInteger resultS))
        {
            return resultS;
        }
        //float
        // ±1.5 x 10−45 to ±3.4 x 1038 	~6-9 digits 	4 bytes

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
    public object? RawValue { get; init; }

    /// <InheritDoc />
    public CellType RawExcelType { get; init; }

    /// <InheritDoc />
    public ReadOnlyMemory<char> ColumnLetters { get; init; }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; init; }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static CellType GetCellType(char[] b, int l)
    {
        switch (b[0])
        {
            case 'b':
                return CellType.Boolean;
            case 'e':
                return CellType.Error;
            case 's':
                return l == 1 ? CellType.SharedString : CellType.String;
            case 'i':
                return CellType.InlineString;
            case 'd':
                return CellType.Date;
            case 'n':
                return CellType.Numeric;
            default:
                // TODO:
                throw new InvalidDataException();
        }
    }

}

