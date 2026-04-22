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

[DebuggerDisplay("{ToString(),raw}")]
internal sealed record Cell : ICell
{
    private static readonly char[]?[] s_columnLetterCache = new char[256][];
    private char[]? _columnLetters;

    // CHANGED: Removed AggressiveOptimization - large method with complex branching, let JIT tier appropriately
    public static async Task<Cell?> ConstructCellAsync(XmlReader reader, InstanceContext instanceContext,
        ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
        CellValue? value = null;
        int col = -1;
        int bufferSize = buffer.Length;
        int len;
        int style = -1;
        bool noCellConversion = instanceContext.Options.CellConversionType <= CellConversion.None;
        bool noCellStyle = instanceContext.Options.CellConversionType < CellConversion.ExcelCellStyle;
        bool returnDBNull = instanceContext.Options.ReturnDBNull;

        void ReadValue()
        {
            len = reader.ReadValueChunk(buffer, 0, bufferSize);
        }

        int expectedAttributes = noCellStyle ? 2 : 3;
        while (reader.MoveToNextAttribute() && expectedAttributes > 0)
        {
            // Retrieve the atomized name directly.
            string currentAttributeName = reader.LocalName;
            if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                ReadValue();
                col = ExcelColumns.ParseColumnOffset(buffer, len);
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                ReadValue();
                type = GetCellType(buffer, len);
                expectedAttributes--;
            }
            else if (!noCellStyle
                && ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom)
                )
            {
                ReadValue();
                style = GetStyleOffset(buffer, len);
                expectedAttributes--;
            }
        }

        reader.MoveToElement();
        if (!reader.IsEmptyElement
            && reader.ReadToDescendant(readerAtoms.vRefAtom)
           )
        {
            // Move to data
            await reader.ReadAsync().ConfigureAwait(false);
            if (noCellConversion)
            {
                if (type == CellType.SharedString)
                {
                    ReadValue();
                    if (len == 0
                        && returnDBNull)
                    {
                        value = new CellValue(DBNull.Value, style);
                    }
                    else
                    {
                        value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()], style);
                    }
                }
                else
                {
                    string? str = ReadString(reader, valueBuilder, buffer);
                    value = returnDBNull
                            && string.IsNullOrEmpty(str)
                        ? new CellValue(DBNull.Value, style)
                        : new CellValue(str, style);
                }
            }
            else //if (instanceContext.Options.CellConversionType >= CellConversion.ExcelCellType)
            {   // Perform conversion
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        {
                            string? str = ReadString(reader, valueBuilder, buffer);
                            value = returnDBNull
                                    && string.IsNullOrEmpty(str)
                                ? new CellValue(DBNull.Value, style)
                                : new CellValue(str, style);
                        }
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            ReadOnlySpan<char> numericSpan = buffer.AsSpan(0, len);
                            value = double.TryParse(numericSpan, NumberStyles.Float, CultureInfo.InvariantCulture, out double numericValue)
                                ? new CellValue(numericValue, style)
                                // If numeric parsing fails, treat as string but avoid intermediate allocation
                                : new CellValue(numericSpan.ToString(), style);
                        }
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()],
                                style);
                        }

                        break;

                    case CellType.Boolean:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            value = new CellValue(buffer[0] != '0');
                        }

                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = new CellValue(ReadString(reader, valueBuilder, buffer), style);
                        break;

                    case CellType.Date:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            ReadOnlySpan<char> dateSpan = buffer.AsSpan(0, len);
                            if (double.TryParse(dateSpan, NumberStyles.Number, CultureInfo.InvariantCulture, out double dateTimeValue))
                            {
                                value = new CellValue(DateTime.FromOADate(dateTimeValue), style);
                            }
                            else if (DateTime.TryParse(dateSpan, out DateTime result))
                            {
                                value = new CellValue(result, style);
                            }
                            else
                            {
                                // If date parsing fails, treat as string; create once from buffer (not dateTimeValue!)
                                value = new CellValue(dateSpan.ToString(), style);
                            }
                        }
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return value is null
            ? null    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
            : new Cell
            {
                //RowNumber = row;
                ExcelColumnOffset = col,
                RawExcelType = type,
                CellValue = value
            };
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    public static Cell? ConstructCell(XmlReader reader, InstanceContext instanceContext,
    ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
        CellValue? value = null;
        int col = -1;
        int bufferSize = buffer.Length;
        int len;
        int style = -1;
        bool noCellConversion = instanceContext.Options.CellConversionType <= CellConversion.None;
        bool noCellStyle = instanceContext.Options.CellConversionType < CellConversion.ExcelCellStyle;
        bool returnDBNull = instanceContext.Options.ReturnDBNull;

        void ReadValue()
        {
            len = reader.ReadValueChunk(buffer, 0, bufferSize);
        }

        int expectedAttributes = noCellStyle ? 2 : 3;
        while (reader.MoveToNextAttribute() && expectedAttributes > 0)
        {
            // Retrieve the atomized name directly.
            string currentAttributeName = reader.LocalName;
            if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                ReadValue();
                col = ExcelColumns.ParseColumnOffset(buffer, len);
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                ReadValue();
                type = GetCellType(buffer, len);
                expectedAttributes--;
            }
            else if (!noCellStyle
                     && ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom)
                    )
            {
                ReadValue();
                style = GetStyleOffset(buffer, len);
                expectedAttributes--;
            }
        }

        reader.MoveToElement();
        if (!reader.IsEmptyElement
            && reader.ReadToDescendant(readerAtoms.vRefAtom)
           )
        {
            // Move to data
            reader.Read();
            if (noCellConversion)
            {
                if (type == CellType.SharedString)
                {
                    ReadValue();
                    if (len == 0
                        && returnDBNull)
                    {
                        value = new CellValue(DBNull.Value, style);
                    }
                    else
                    {
                        value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()], style);
                    }
                }
                else
                {
                    string? str = ReadString(reader, valueBuilder, buffer);
                    value = returnDBNull
                            && string.IsNullOrEmpty(str)
                        ? new CellValue(DBNull.Value, style)
                        : new CellValue(str, style);
                }
            }
            else //if (instanceContext.Options.CellConversionType >= CellConversion.ExcelCellType)
            {   // Perform conversion
                switch (type)
                {
                    case CellType.Unknown:
                    case CellType.Formula:
                    case CellType.InlineString:
                        {
                            string? str = ReadString(reader, valueBuilder, buffer);
                            value = returnDBNull
                                    && string.IsNullOrEmpty(str)
                                ? new CellValue(DBNull.Value, style)
                                : new CellValue(str, style);
                        }
                        break;

                    case CellType.Numeric:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            ReadOnlySpan<char> numericSpan = buffer.AsSpan(0, len);
                            value = double.TryParse(numericSpan, NumberStyles.Float, CultureInfo.InvariantCulture,
                                out double numericValue)
                                ? new CellValue(numericValue, style)
                                // If numeric parsing fails, treat as string but avoid intermediate allocation
                                : new CellValue(numericSpan.ToString(), style);
                        }
                        break;

                    case CellType.SharedString:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            value = new CellValue(instanceContext?.SharedStrings?[buffer.AsSpan(0, len).IntParse()],
                                style);
                        }
                        break;

                    case CellType.Boolean:
                        ReadValue();
                        if (len == 0
                            && returnDBNull)
                        {
                            value = new CellValue(DBNull.Value, style);
                        }
                        else
                        {
                            value = new CellValue(buffer[0] != '0');
                        }
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = new CellValue(ReadString(reader, valueBuilder, buffer), style);
                        break;

                    case CellType.Date:
                        {
                            ReadValue();
                            if (len == 0
                                && returnDBNull)
                            {
                                value = new CellValue(DBNull.Value, style);
                                break;
                            }

                            ReadOnlySpan<char> dateSpan = buffer.AsSpan(0, len);
                            if (double.TryParse(dateSpan, NumberStyles.Number, CultureInfo.InvariantCulture,
                                    out double dateTimeValue))
                            {
                                value = new CellValue(DateTime.FromOADate(dateTimeValue), style);
                            }
                            else if (DateTime.TryParse(dateSpan, out DateTime result))
                            {
                                value = new CellValue(result, style);
                            }
                            else
                            {
                                // If date parsing fails, treat as string; create once from buffer (not dateTimeValue!)
                                value = new CellValue(dateSpan.ToString(), style);
                            }
                        }
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return value is null
            ? null    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
            : new Cell
            {
                //RowNumber = row;
                ExcelColumnOffset = col,
                RawExcelType = type,
                CellValue = value
            };
    }

    #region Borrowed and some finessing from XMLReader source
    // CHANGED: Removed AggressiveOptimization - complex method with loops and branches
    private static string? ReadString(XmlReader reader, StringBuilder valueBuilder, char[] buffer)
    {
        if (reader.ReadState != ReadState.Interactive)
        {
            return null;
        }

        // If we're positioned on an element, parse its inner textual content by
        // using ReadElementContentAsString fast path where possible, otherwise
        // collect text with ReadValueChunk into the provided buffer to minimize allocations.
        if (reader.NodeType == XmlNodeType.Element)
        {
            if (reader.IsEmptyElement)
            {
                return null;
            }

            int startDepth = reader.Depth;

            // Move to the first node inside the element
            if (!reader.Read())
            {
                return null;
            }

            valueBuilder.Length = 0;
            bool readAny = false;

            // Use chunked reading to avoid extra allocations
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
                            readAny = true;
                        }
                    } while (read > 0);
                }

                if (!reader.Read())
                {
                    break;
                }
            }

            return readAny ? valueBuilder.ToString() : string.Empty;
        }

        // Not positioned on an element - return value if textual
        XmlNodeType nodeType = reader.NodeType;
        return nodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace
            ? reader.Value
            : string.Empty;
    }

    #endregion

    /// <InheritDoc />
    public CellValue? CellValue { get; internal init; }

    /// <InheritDoc />
    public CellType RawExcelType { get; private init; }

    /// <InheritDoc />
    public IReadOnlyList<char> ColumnLetters
    {
        // CHANGED: Removed AggressiveOptimization - simple property with cache lookup, inline better
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get
        {
            if (_columnLetters != null)
            {
                return _columnLetters;
            }

            int offset = ExcelColumnOffset;
            if (offset <= 0)
            {
                return _columnLetters = Array.Empty<char>();
            }

            if (offset < s_columnLetterCache.Length)
            {
                char[]? cached = s_columnLetterCache[offset];
                if (cached != null)
                {
                    return _columnLetters = cached;
                }

                char[] computed = offset.GetExcelColumnName().ToCharArray();
                s_columnLetterCache[offset] = computed;
                return _columnLetters = computed;
            }

            return _columnLetters = offset.GetExcelColumnName().ToCharArray();
        }
    }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; internal init; }

    /// <InheritDoc />
    public override string? ToString() => CellValue?.ToString() ?? base.ToString();

    // CHANGED: Removed AggressiveOptimization - simple switch on first char, inline better
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
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
            's' => l == 1 ? CellType.SharedString : /*"str"*/CellType.Formula,
            'i' => /*"inlineStr"*/CellType.InlineString,
            'd' => CellType.Date,
            'n' => CellType.Numeric,
            _ => throw new InvalidDataException()
        };
    }

    private static int GetStyleOffset(char[] b, int l) => l == 0 ? -1 : b.AsSpan(0, l).IntParse();
}