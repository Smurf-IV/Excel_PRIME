using System.Diagnostics;
using System.Runtime.CompilerServices;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{ToString(),raw}")]
internal sealed record XlsbCell : ICell
{
    // Static cache for column letter arrays to avoid repeated allocations
    // Column offsets range from 1-16384 in Excel, allocate conservatively
    private static readonly char[]?[] s_columnLetterCache = new char[256][];

    // CHANGED: Removed AggressiveOptimization - switch-based method, let JIT optimize naturally
    public static XlsbCell? ConstructCell(PooledRecordBuffer reader, InstanceContext instanceContext)
    {
        int columnOffset = reader.GetInt32(0) + 1; // Convert zero-based to Excel one-based
        short styleRef = (short)(instanceContext.Options.CellConversionType <= CellConversion.None ? -1 : reader.GetInt32(4));

        CellType cellType;
        CellValue? cellValue;
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
                return null;
        }

        // Allocate once and populate properties
        return new XlsbCell
        {
            ExcelColumnOffset = columnOffset,
            RawExcelType = cellType,
            CellValue = cellValue
        };
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static string? GetSharedString(InstanceContext instanceContext, PooledRecordBuffer reader) => instanceContext.SharedStrings?[reader.GetInt32(8)];

    // CHANGED: Kept AggressiveInlining, removed AggressiveOptimization - hot-path bit manipulation method benefits from inline
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
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

    /// <InheritDoc />
    public CellValue? CellValue { get; internal init; }

    /// <InheritDoc />
    public CellType RawExcelType { get; private init; }

    /// <InheritDoc />
    public IReadOnlyList<char> ColumnLetters
    {
        // CHANGED: Removed AggressiveOptimization - simple cache lookup property, inline better
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get
        {
            int offset = ExcelColumnOffset;
            if (offset <= 0)
            {
                return [];
            }
            // Check cache first for common column ranges
            if (offset < s_columnLetterCache.Length)
            {
                char[]? cached = s_columnLetterCache[offset];
                if (cached != null)
                {
                    return cached;
                }
                // Compute and cache the column letters
                char[] result = offset.GetExcelColumnName();
                s_columnLetterCache[offset] = result;
                return result;
            }
            // Fallback for out-of-range offsets
            return offset.GetExcelColumnName();
        }
    }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; internal init; }

    /// <InheritDoc />
    public override string? ToString() => CellValue?.ToString() ?? base.ToString();

}
