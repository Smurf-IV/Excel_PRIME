using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{CellValue.ToString(),raw}")]
internal sealed record Cell : ICell
{
    public static async Task<Cell?> ConstructCellAsync(XmlReader reader, InstanceContext instanceContext,
        ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
        CellValue? value = null;
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
                    value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()]);
                }
                else
                {
                    value = new CellValue(ReadString(reader, valueBuilder, buffer));
                }
            }
            else
            {
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        value = new CellValue(ReadString(reader, valueBuilder, buffer));
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        value = new CellValue(double.Parse(buffer.AsSpan(0, len), NumberStyles.Float,
                            CultureInfo.InvariantCulture));
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()]);
                        break;

                    case CellType.Boolean:
                        ReadValue();
                        value = new CellValue(buffer[0] != '0');
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = new CellValue(ReadString(reader, valueBuilder, buffer));
                        break;

                    case CellType.Date:
                        ReadValue();
                        if (double.TryParse(buffer.AsSpan(0, len), NumberStyles.Number,
                                CultureInfo.InvariantCulture, out double dateTimeValue))
                        {
                            value = new CellValue(DateTime.FromOADate(dateTimeValue));
                        }
                        else if (DateTime.TryParse(buffer.AsSpan(0, len), out DateTime result))
                        {
                            value = new CellValue(result);
                        }
                        else
                        {
                            value = new CellValue(new string(buffer, 0, len));
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
                CellValue = value.Value
            };
    }

    public static Cell? ConstructCell(XmlReader reader, InstanceContext instanceContext,
    ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
        CellValue? value = null;
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
                    value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()]);
                }
                else
                {
                    value = new CellValue(ReadString(reader, valueBuilder, buffer));
                }
            }
            else
            {
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        value = new CellValue(ReadString(reader, valueBuilder, buffer));
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        value = new CellValue(double.Parse(buffer.AsSpan(0, len), NumberStyles.Float,
                            CultureInfo.InvariantCulture));
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()]);
                        break;

                    case CellType.Boolean:
                        ReadValue();
                        value = new CellValue(buffer[0] != '0');
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = new CellValue(ReadString(reader, valueBuilder, buffer));
                        break;

                    case CellType.Date:
                        ReadValue();
                        if (double.TryParse(buffer.AsSpan(0, len), NumberStyles.Number,
                                CultureInfo.InvariantCulture, out double dateTimeValue))
                        {
                            value = new CellValue(DateTime.FromOADate(dateTimeValue));
                        }
                        else if (DateTime.TryParse(buffer.AsSpan(0, len), out DateTime result))
                        {
                            value = new CellValue(result);
                        }
                        else
                        {
                            value = new CellValue(new string(buffer,0, len));
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
                CellValue = value.Value
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
        return nodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace
            ? reader.Value
            : string.Empty;
    }

    #endregion
    
    /// <InheritDoc />
    public CellValue CellValue { get; private init; }

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