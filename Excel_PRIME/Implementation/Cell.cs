using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;
using ExcelPRIME.XlsbImp;


namespace ExcelPRIME;

[DebuggerDisplay("{ToString(),raw}")]
[StructLayout(LayoutKind.Explicit, Size = 32)]
public readonly struct Cell : ICell
{
    [FieldOffset(0)]
    private readonly int _packedInfo;

    [FieldOffset(4)]
    private readonly CellValue _cellValue;

    /// <InheritDoc />
    public CellValue CellValue => _cellValue;

    /// <InheritDoc />
    public CellType RawExcelType => (CellType)(_packedInfo >> 24);

    /// <InheritDoc />
    public int ExcelColumnOffset => _packedInfo & 0x00FFFFFF;

    internal Cell(CellValue value, int col, CellType type)
    {
        _cellValue = value;
        _packedInfo = (col & 0x00FFFFFF) | ((int)type << 24);
    }

    private static readonly char[]?[] s_columnLetterCache = new char[256][];

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal static async ValueTask<Cell> ConstructCellAsync(XmlReader reader, InstanceContext instanceContext,
        ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
        CellValue value = default;
        int col = -1;
        int bufferSize = buffer.Length;
        int len;
        short style = -1;
        bool noCellConversion = instanceContext.Options.CellConversionType <= CellConversion.None;
        bool returnDBNull = instanceContext.Options.ReturnDBNull;

        int expectedAttributes = noCellConversion ? 2 : 3;
        while (reader.MoveToNextAttribute() && expectedAttributes > 0)
        {
            // Retrieve the atomized name directly.
            string currentAttributeName = reader.LocalName;
            if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                len = ReadValue(reader, buffer, bufferSize);
                col = ExcelColumns.ParseColumnOffset(buffer.AsSpan(0, len));
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                len = ReadValue(reader, buffer, bufferSize);
                type = GetCellType(buffer, len);
                expectedAttributes--;
            }
            else if (!noCellConversion
                && ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom)
                )
            {
                len = ReadValue(reader, buffer, bufferSize);
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
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                // Handle empty value "EndElement" cell, e.g. <c r="F7"/>
                goto setter;
            }
            if (noCellConversion)
            {
                switch (type)
                {
                    case CellType.SharedString:
                        len = ReadValue(reader, buffer, bufferSize);
                        if (len == 0)
                        {
                            goto setter;
                        }

                        value = CellValue.Create(instanceContext.SharedStrings?[buffer.AsSpan(0, len).IntParse()], style);
                        break;

                    case CellType.Date:
                        value = ConvertToDate(reader, buffer, returnDBNull, style);
                        break;

                    default:
                        {
                            string? str = ReadString(reader, valueBuilder, buffer);
                            if (returnDBNull
                                && string.IsNullOrEmpty(str))
                            {
                                goto setter;
                            }

                            value = CellValue.Create(str, style);
                        }
                        break;
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
                            if (returnDBNull
                                && string.IsNullOrEmpty(str))
                            {
                                goto setter;
                            }

                            value = CellValue.Create(str, style);
                        }
                        break;

                    case CellType.Numeric:
                        {
                            len = ReadValue(reader, buffer, bufferSize);
                            if (len == 0)
                            {
                                goto setter;
                            }
                            CellStyle? cellStyle = null;
                            if (instanceContext.CellStyles?.TryGetValue(style, out cellStyle) == false)
                            {
                                cellStyle = null;
                            }

                            value = CellValue.TryParseOrder(buffer.AsSpan(0, len), cellStyle);
                        }
                        break;

                    case CellType.SharedString:
                        len = ReadValue(reader, buffer, bufferSize);
                        if (len == 0)
                        {
                            goto setter;
                        }

                        value = CellValue.Create(instanceContext.SharedStrings?[buffer.AsSpan(0, len).IntParse()],
                            style);

                        break;

                    case CellType.Boolean:
                        len = ReadValue(reader, buffer, bufferSize);
                        if (len == 0)
                        {
                            goto setter;
                        }

                        value = CellValue.Create(buffer[0] != '0');

                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = CellValue.Create(ReadString(reader, valueBuilder, buffer), style);
                        break;

                    case CellType.Date:
                        value = ConvertToDate(reader, buffer, returnDBNull, style);
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        setter:
        if (returnDBNull
            && value.IsUnknown)
        {
            value = CellValue.GetDBNull(style);
        }

        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return value.IsUnknown
            ? default    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
            : new Cell(value, col, type);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal static Cell ConstructXlsbCell(PooledRecordBuffer reader, InstanceContext instanceContext)
    {
        int columnOffset = reader.GetInt32(0) + 1; // Convert zero-based to Excel one-based
        short styleRef = (short)(instanceContext.Options.CellConversionType <= CellConversion.None ? -1 : reader.GetInt32(4));

        CellType cellType;
        CellValue cellValue = default;
        switch (reader.RecordType)
        {
            case RecordTypeIdentifier.CELLRK:
                (cellType, cellValue) = MagicConvertRK(reader, styleRef, instanceContext);
                break;
            case RecordTypeIdentifier.CELLREAL or RecordTypeIdentifier.CELLFMLANUM:
                {
                    double d = reader.GetDouble(8);
                    if ((instanceContext.CellStyles?.TryGetValue(styleRef, out CellStyle? cellStyle) == true)
                        && cellStyle?.IsDateStyle == true
                       )
                    {
                        (cellType, cellValue) = (CellType.Date, CellValue.Create(DateTime.FromOADate(d), cellStyle.ExcelFormatId));
                    }
                    else
                    {
                        (cellType, cellValue) = (CellType.Numeric, CellValue.Create(d, styleRef));
                    }
                }
                break;
            case RecordTypeIdentifier.CELLBOOL or RecordTypeIdentifier.CELLFMLABOOL:
                (cellType, cellValue) = (CellType.Boolean, CellValue.Create(reader.GetByte(8) != 0));
                break;
            case RecordTypeIdentifier.CELLST or RecordTypeIdentifier.CELLFMLASTRING:
                (cellType, cellValue) = (CellType.InlineString, CellValue.Create(reader.GetString(8), styleRef));
                break;
            case RecordTypeIdentifier.CELLISST:
                (cellType, cellValue) = (CellType.SharedString,
                    CellValue.Create(GetSharedString(instanceContext, reader), styleRef));
                break;
            case RecordTypeIdentifier.CELLERROR or RecordTypeIdentifier.CELLFMLAERROR:
                (cellType, cellValue) = (CellType.Error, CellValue.Create((ExcelErrorCode)reader.GetByte(8)));
                break;
            default:
                // Break out early
                return default;
        }

        // Allocate once and populate properties
        return new Cell(cellValue, columnOffset, cellType);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static string? GetSharedString(InstanceContext instanceContext, PooledRecordBuffer reader) => instanceContext.SharedStrings?[reader.GetInt32(8)];

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    private static (CellType, CellValue) MagicConvertRK(PooledRecordBuffer record, short styleRef, InstanceContext instanceContext)
    {
        // Extract and process the RK value in a single optimized path
        int rk = record.GetInt32(8);

        double d;

        if ((rk & 0x02) == 0) // isFloat
        {
            // Float encoding: mask off the type bits and shift to 64-bit representation
            long v = rk & 0xfffffffc;
            v <<= 32;
            d = BitConverter.Int64BitsToDouble(v);
        }
        else
        {
            // Integer encoding: shift right by 2 to remove type bits
            d = rk >> 2;
        }

        // Check if scaled by 100
        if ((rk & 0x01) != 0)
        {
            d /= 100.0;  // Explicit double to ensure double division
        }

        if ((instanceContext.CellStyles?.TryGetValue(styleRef, out CellStyle? cellStyle) == true)
            && cellStyle?.IsDateStyle == true
            )
        {
            return (CellType.Date, CellValue.Create(DateTime.FromOADate(d), cellStyle.ExcelFormatId));
        }

        return double.IsInteger(d)
            ? (CellType.Numeric, CellValue.Create((int)d, styleRef))
            : (CellType.Numeric, CellValue.Create(d, styleRef));
    }

    private static int ReadValue(XmlReader reader, char[] buffer, int bufferSize)
        => reader.ReadValueChunk(buffer, 0, bufferSize);

    private static CellValue ConvertToDate(XmlReader reader, char[] buffer, bool returnDBNull, short style)
    {
        CellValue value;
        int len = ReadValue(reader, buffer, buffer.Length);
        if (len == 0
            && returnDBNull)
        {
            value = CellValue.GetDBNull(style);
        }
        else
        {
            ReadOnlySpan<char> dateSpan = buffer.AsSpan(0, len);
            if (dateSpan.TryDoubleParse(out double dateTimeValue))
            {
                value = CellValue.Create(DateTime.FromOADate(dateTimeValue), style);
            }
            else if (DateTime.TryParse(dateSpan, out DateTime result))
            {
                value = CellValue.Create(result, style);
            }
            else
            {
                // If date parsing fails, treat as string; create once from buffer (not dateTimeValue!)
                value = CellValue.Create(dateSpan.ToString(), style);
            }
        }

        return value;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    internal static Cell ConstructCell(XmlReader reader, InstanceContext instanceContext,
    ReaderAtoms readerAtoms, char[] buffer, StringBuilder valueBuilder)
    {
        CellType type = CellType.Numeric;
        CellValue value = default;
        int col = -1;
        int bufferSize = buffer.Length;
        int len;
        short style = -1;
        bool noCellConversion = instanceContext.Options.CellConversionType <= CellConversion.None;
        bool returnDBNull = instanceContext.Options.ReturnDBNull;

        int expectedAttributes = noCellConversion ? 2 : 3;
        while (reader.MoveToNextAttribute() && expectedAttributes > 0)
        {
            // Retrieve the atomized name directly.
            string currentAttributeName = reader.LocalName;
            if (ReferenceEquals(currentAttributeName, readerAtoms.rRefAtom))
            {
                len = ReadValue(reader, buffer, bufferSize);
                col = ExcelColumns.ParseColumnOffset(buffer.AsSpan(0, len));
                expectedAttributes--;
            }
            else if (ReferenceEquals(currentAttributeName, readerAtoms.tRefAtom))
            {
                len = ReadValue(reader, buffer, bufferSize);
                type = GetCellType(buffer, len);
                expectedAttributes--;
            }
            else if (!noCellConversion
                     && ReferenceEquals(currentAttributeName, readerAtoms.sRefAtom)
                    )
            {
                len = ReadValue(reader, buffer, bufferSize);
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
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                // Handle empty value "EndElement" cell, e.g. <c r="F7"/>
                goto setter;
            }
            if (noCellConversion)
            {
                switch (type)
                {
                    case CellType.SharedString:
                        len = ReadValue(reader, buffer, bufferSize);
                        if (len == 0)
                        {
                            goto setter;
                        }

                        value = CellValue.Create(instanceContext.SharedStrings?[buffer.AsSpan(0, len).IntParse()], style);
                        break;

                    case CellType.Date:
                        value = ConvertToDate(reader, buffer, returnDBNull, style);
                        break;

                    default:
                        {
                            string? str = ReadString(reader, valueBuilder, buffer);
                            if (returnDBNull
                                && string.IsNullOrEmpty(str))
                            {
                                goto setter;
                            }

                            value = CellValue.Create(str, style);
                        }
                        break;
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
                            if (returnDBNull
                                && string.IsNullOrEmpty(str))
                            {
                                goto setter;
                            }

                            value = CellValue.Create(str, style);
                        }
                        break;

                    case CellType.Numeric:
                        {
                            len = ReadValue(reader, buffer, bufferSize);
                            if (len == 0)
                            {
                                goto setter;
                            }

                            CellStyle? cellStyle = null;
                            if ( !instanceContext?.CellStyles?.TryGetValue(style, out cellStyle) ?? true)
                            {
                                cellStyle = null;
                            }

                            value = CellValue.TryParseOrder(buffer.AsSpan(0, len), cellStyle);
                        }
                        break;

                    case CellType.SharedString:
                        len = ReadValue(reader, buffer, bufferSize);
                        if (len == 0)
                        {
                            goto setter;
                        }
                        value = CellValue.Create(instanceContext.SharedStrings?[buffer.AsSpan(0, len).IntParse()], style);
                        break;

                    case CellType.Boolean:
                        len = ReadValue(reader, buffer, bufferSize);

                        if (len == 0)
                        {
                            goto setter;
                        }

                        value = CellValue.Create(buffer[0] != '0');
                        break;

                    case CellType.Error:
                        // TODO: Decrypt the error
                        value = CellValue.Create(ReadString(reader, valueBuilder, buffer), style);
                        break;

                    case CellType.Date:
                        value = ConvertToDate(reader, buffer, returnDBNull, style);
                        break;

                    default:
                        throw new ArgumentOutOfRangeException(((int)type).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        setter:
        if (returnDBNull
            && value.IsUnknown)
        {
            value = CellValue.GetDBNull(style);
        }

        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        return value.IsUnknown
                ? default    // Deal with an empty value "EndElement" cell, e.g. <c r="B1" s="2" />
                : new Cell(value, col, type);
    }

    #region Borrowed and some finessing from XMLReader source
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
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
    public IReadOnlyList<char> ColumnLetters
    {
        // CHANGED: Removed AggressiveOptimization - simple property with cache lookup, inline better
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get
        {
            int offset = ExcelColumnOffset;
            if (offset <= 0)
            {
                return Array.Empty<char>();
            }

            if (offset < s_columnLetterCache.Length)
            {
                char[]? cached = s_columnLetterCache[offset];
                if (cached != null)
                {
                    return cached;
                }

                char[] computed = offset.GetExcelColumnName();
                s_columnLetterCache[offset] = computed;
                return computed;
            }

            return offset.GetExcelColumnName();
        }
    }

    /// <InheritDoc />
    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    public override string? ToString() => _cellValue.IsUnknown ? string.Empty : _cellValue.ToString();

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
            'f' => CellType.Formula,
            'i' => /*"inlineStr"*/CellType.InlineString,
            'd' => CellType.Date,
            'n' => CellType.Numeric,
            _ => throw new InvalidDataException()
        };
    }

    private static short GetStyleOffset(char[] b, int l) => (short)(l == 0 ? -1 : b.AsSpan(0, l).IntParse());

}
