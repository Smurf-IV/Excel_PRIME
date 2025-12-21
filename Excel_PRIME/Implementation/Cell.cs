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

using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{RawValue.ToString(),raw}")]
internal sealed record Cell : ICell
{
    public static async Task<Cell?> ConstructCellAsync(XmlReader reader, InstanceContext instanceContext,
        ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
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
            if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                ReadValue();
                (int _, col, colName) = new ReadOnlySpan<char>(buffer, 0, len).GetRowColNumbers();
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                ReadValue();
                type = GetCellType(buffer, len);
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom))
            {
                // TODO: the style, therefore converting into time only etc.
                //ReadValue();
                //style = GetStyleOffset(buffer, len);
            }
        }

        reader.MoveToElement();
        if (!reader.IsEmptyElement
            && reader.ReadToDescendant(readerAtoms.vRefAtom)
           )
        {
            // Move to data
            await reader.ReadAsync().ConfigureAwait(false);
            if (instanceContext.Options.CellConversionType == CellConversion.None)
            {
                if (type == CellType.SharedString)
                {
                    ReadValue();
                    value = instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()];
                }
                else
                {
                    value = ReadString(reader, valueBuilder, buffer);
                }
            }
            else
            {
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        value = ReadString(reader, valueBuilder, buffer);
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        value = TryParseOrder(instanceContext.Options.CellConversionType, buffer.AsSpan(0, len));
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        value = instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()];
                        break;

                    case CellType.Boolean:
                        ReadValue();
                        value = buffer[0] != '0';
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = ReadString(reader, valueBuilder, buffer);
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
                            // Prefer returning ReadOnlyMemory<char> to avoid allocating a new string
                            value = new ReadOnlyMemory<char>(buffer.AsSpan(0, len).ToArray());
                        }
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return value == null
            ? null    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
            : new Cell
            {
                ColumnLetters = colName,
                //RowNumber = row;
                ExcelColumnOffset = col,
                RawExcelType = type,
                RawValue = value
            };
    }

    public static Cell? ConstructCell(XmlReader reader, InstanceContext instanceContext,
    ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
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
            if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                ReadValue();
                (int _, col, colName) = new ReadOnlySpan<char>(buffer, 0, len).GetRowColNumbers();
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                ReadValue();
                type = GetCellType(buffer, len);
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom))
            {
                // TODO: the style, therefore converting into time only etc.
                //ReadValue();
                //style = GetStyleOffset(buffer, len);
            }
        }

        reader.MoveToElement();
        if (!reader.IsEmptyElement
            && reader.ReadToDescendant(readerAtoms.vRefAtom)
           )
        {
            // Move to data
            reader.Read();
            if (instanceContext.Options.CellConversionType == CellConversion.None)
            {
                if (type == CellType.SharedString)
                {
                    ReadValue();
                    value = instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()];
                }
                else
                {
                    value = ReadString(reader, valueBuilder, buffer);
                }
            }
            else
            {
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        value = ReadString(reader, valueBuilder, buffer);
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        value = TryParseOrder(instanceContext.Options.CellConversionType, buffer.AsSpan(0, len));
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        value = instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()];
                        break;

                    case CellType.Boolean:
                        ReadValue();
                        value = buffer[0] != '0';
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = ReadString(reader, valueBuilder, buffer);
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
                            value = new ReadOnlyMemory<char>(buffer.AsSpan(0, len).ToArray());
                        }
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return value == null
            ? null    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
            : new Cell
            {
                ColumnLetters = colName,
                //RowNumber = row;
                ExcelColumnOffset = col,
                RawExcelType = type,
                RawValue = value
            };
    }

    #region Borrowed and some finessing from XMLReader source
    private static string ReadString(XmlReader reader, StringBuilder valueBuilder, char[] buffer)
    {
        if (reader.ReadState != ReadState.Interactive)
        {
            return string.Empty;
        }

        // If we're positioned on an element, parse its inner textual content by
        // using ReadElementContentAsString fast path where possible, otherwise
        // collect text with ReadValueChunk into the provided buffer to minimize allocations.
        if (reader.NodeType == XmlNodeType.Element)
        {
            if (reader.IsEmptyElement)
            {
                return string.Empty;
            }

            // Try the optimized fast path first
            try
            {
                return reader.ReadElementContentAsString();
            }
            catch
            {
                // fallthrough to manual chunked read
            }

            int startDepth = reader.Depth;

            // Move to the first node inside the element
            if (!reader.Read())
            {
                return string.Empty;
            }

            valueBuilder.Clear();
            bool wroteAny = false;

            while (reader.Depth > startDepth)
            {
                XmlNodeType nt = reader.NodeType;
                bool isTextual = nt is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace;
                if (isTextual)
                {
                    // Read text content in chunks into buffer and append
                    int read;
                    do
                    {
                        read = reader.ReadValueChunk(buffer, 0, buffer.Length);
                        if (read > 0)
                        {
                            valueBuilder.Append(buffer, 0, read);
                            wroteAny = true;
                        }
                    } while (read > 0);
                }

                if (!reader.Read())
                {
                    break;
                }
            }

            return wroteAny ? valueBuilder.ToString() : string.Empty;
        }

        // Not positioned on an element - return value if textual
        XmlNodeType nodeType = reader.NodeType;
        if (nodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace)
        {
            return reader.Value;
        }

        return string.Empty;
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

    private static object? TryParseOrder(CellConversion optionsCell_conversionType, ReadOnlySpan<char> asSpan)
    {
        if (asSpan.Length == 0)
        {
            return null;
        }

        return optionsCell_conversionType switch
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
        {   //   	±5.0 × 10−324 to ±1.7 × 10308 	~15-17 digits 	8 bytes
            return resultD;
        }
        // Fall back to ReadOnlyMemory to avoid immediate string allocation; callers that need string can ToString()
        return new ReadOnlyMemory<char>(asSpan.ToArray());
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