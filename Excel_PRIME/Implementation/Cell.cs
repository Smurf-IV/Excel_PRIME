using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{RawValue.ToString(),raw}")]
internal class Cell : ICell
{
    public static async Task<Cell> ConstructCellAsync(XmlReader reader, InstanceContext instanceContext,
        ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Unknown;
        object? value = null;
        int col = -1;
        char[] colName = [];
        int bufferSize = buffer.Length;
        int len;

        void ReadValue()
        {
            len = reader.ReadValueChunk(buffer, 0, bufferSize);
        }

        int expectedAttributes = 2;
        while (reader.MoveToNextAttribute() && expectedAttributes > 0)
        {
            // Retrieve the atomized name directly.
            string currentAttributeName = reader.LocalName;
            if (Object.ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                ReadValue();
                (int _, col, colName) = new ReadOnlySpan<char>(buffer, 0, len).GetRowColNumbers();
                expectedAttributes--;
            }
            else if (Object.ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                ReadValue();
                type = GetCellType(buffer, len);
                expectedAttributes--;
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
            if (instanceContext.Options.CellConversionType == CellConversion.None)
            {
                if (type == CellType.SharedString)
                {
                    ReadValue();
                    value = instanceContext.SharedStrings[buffer.AsSpan(0, len).IntParse()];
                }
                else
                {
                    value = ReadString(reader, valueBuilder);
                }
            }
            else
            {
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        value = ReadString(reader, valueBuilder);
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        value = TryParseOrder(instanceContext.Options.CellConversionType, buffer.AsSpan(0, len));
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        value = instanceContext.SharedStrings[buffer.AsSpan(0, len).IntParse()];
                        break;

                    case CellType.Boolean:
                        ReadValue();
                        value = buffer[0] != '0';
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = ReadString(reader, valueBuilder);
                        break;

                    case CellType.Date:
                        ReadValue();
                        if (double.TryParse(buffer.AsSpan(0, len), NumberStyles.Number,
                                CultureInfo.InvariantCulture, out double dateTimeValue))
                        {
                            value = DateTime.FromOADate(dateTimeValue);
                        }
                        else if (DateTime.TryParse(buffer.AsSpan(0, len), out DateTime result))
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
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return new Cell
        {
            ColumnLetters = colName,
            //RowNumber = row;
            ExcelColumnOffset = col,
            RawExcelType = type,
            RawValue = value
        };
    }

    #region Borrowed and some finessing from XMLReader source
    private static string ReadString(XmlReader reader, StringBuilder valueBuilder)
    {
        if (reader.ReadState != ReadState.Interactive)
        {
            return string.Empty;
        }
        XmlReader subReader = reader;

        if (subReader.NodeType == XmlNodeType.Element)
        {
            if (subReader.IsEmptyElement)
            {
                return string.Empty;
            }

            subReader = reader.ReadSubtree();
            if (subReader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    return string.Empty;
                }
            }
        }
        string result = string.Empty;
        int hasMultipleTextForCell = 0;
        while (IsTextualNode(reader.NodeType))
        {
            if (hasMultipleTextForCell++ > 0)
            {
                valueBuilder.Append(result);
            }
            result = reader.Value;
            if (!subReader.Read())
            {
                break;
            }
        }
        if (hasMultipleTextForCell > 1)
        {
            // Add last iteration, and get current combined string
            valueBuilder.Append(result);
            result = valueBuilder.ToString();
        }
        return result;
    }

    private const uint IsTextualNodeBitmap = 0x6018; // 00 0110 0000 0001 1000
                                                     // 0 None,
                                                     // 0 Element,
                                                     // 0 Attribute,
                                                     // 1 Text,
                                                     // 1 CDATA,
                                                     // 0 EntityReference,
                                                     // 0 Entity,
                                                     // 0 ProcessingInstruction,
                                                     // 0 Comment,
                                                     // 0 Document,
                                                     // 0 DocumentType,
                                                     // 0 DocumentFragment,
                                                     // 0 Notation,
                                                     // 1 Whitespace,
                                                     // 1 SignificantWhitespace,
                                                     // 0 EndElement,
                                                     // 0 EndEntity,
                                                     // 0 XmlDeclaration

    private static bool IsTextualNode(XmlNodeType nodeType)
    {
#if DEBUG
            // This code verifies IsTextualNodeBitmap mapping of XmlNodeType to a bool specifying
            // whether the node is 'textual' = Text, CDATA, Whitespace or SignificantWhitespace.
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.None)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Element)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Attribute)));
            Debug.Assert(0 != (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Text)));
            Debug.Assert(0 != (IsTextualNodeBitmap & (1 << (int)XmlNodeType.CDATA)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.EntityReference)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Entity)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.ProcessingInstruction)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Comment)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Document)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.DocumentType)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.DocumentFragment)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Notation)));
            Debug.Assert(0 != (IsTextualNodeBitmap & (1 << (int)XmlNodeType.Whitespace)));
            Debug.Assert(0 != (IsTextualNodeBitmap & (1 << (int)XmlNodeType.SignificantWhitespace)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.EndElement)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.EndEntity)));
            Debug.Assert(0 == (IsTextualNodeBitmap & (1 << (int)XmlNodeType.XmlDeclaration)));
#endif
        return 0 != (IsTextualNodeBitmap & (1 << (int)nodeType));
    }
    #endregion

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
            CellConversion.ForceStyles => // TODO
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
            // double
            // ±1.5 x 10−45 to ±3.4 x 1038 	~6-9 digits 	4 bytes
            return PerformDoubleConversion(asSpan);
        }

        // ReSharper disable once ConvertIfStatementToSwitchStatement
        if (asSpan.Length < 12
           //&& int.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resultI)
           )
        {
            // -2,147,483,648 to 2,147,483,647 	Signed 32-bit integer
            if (asSpan[0] != '-')
            {
                return asSpan.IntParse();
            }

            if (int.TryParse(asSpan, NumberStyles.Integer, CultureInfo.InvariantCulture, out int resultI))
            {
                return resultI;
            }
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

        return PerformDoubleConversion(asSpan);
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static object PerformDoubleConversion(ReadOnlySpan<char> asSpan)
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
    public IReadOnlyList<char> ColumnLetters { get; private init; } = null!;

    /// <InheritDoc />
    public int ExcelColumnOffset { get; private init; }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static CellType GetCellType(in char[] b, int l)
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
            's' => l == 1 ? CellType.SharedString : /*"str"*/CellType.Formula,
            'i' => /*"inlineStr"*/CellType.InlineString,
            'd' => CellType.Date,
            'n' => CellType.Numeric,
            _ => throw new InvalidDataException()
        };
    }
}