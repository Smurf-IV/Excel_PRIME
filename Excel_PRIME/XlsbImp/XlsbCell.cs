using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.CompilerServices;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{CellValue}")]
internal sealed record XlsbCell : ICell
{
    // Static cache for column letter arrays to avoid repeated allocations
    // Column offsets range from 1-16384 in Excel, allocate conservatively
    private static readonly char[]?[] s_columnLetterCache = new char[256][];

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static XlsbCell? ConstructCell(PooledRecordBuffer reader, InstanceContext instanceContext)
    {
        int columnOffset = reader.GetInt32(0) + 1; // Convert zero-based to one-based
        RecordTypeIdentifier recordType = reader.RecordType;

        // Use expression-based switch for efficient dispatch
        CellType cellType = CellType.Unknown;
        CellValue? cellValue = null;
        switch (recordType)
        {
            case RecordTypeIdentifier.CELLRK:
                (cellType, cellValue) = (CellType.Numeric, new CellValue(MagicConvertRK(reader)));
                break;
            case RecordTypeIdentifier.CELLREAL or RecordTypeIdentifier.CELLFMLANUM:
                (cellType, cellValue) = (CellType.Numeric, new CellValue(reader.GetDouble(8)));
                break;
            case RecordTypeIdentifier.CELLBOOL or RecordTypeIdentifier.CELLFMLABOOL:
                (cellType, cellValue) = (CellType.Boolean, new CellValue(reader.GetByte(8) != 0));
                break;
            case RecordTypeIdentifier.CELLST or RecordTypeIdentifier.CELLFMLASTRING:
                (cellType, cellValue) = (CellType.InlineString, new CellValue(reader.GetString(8)));
                break;
            case RecordTypeIdentifier.CELLISST:
                (cellType, cellValue) = (CellType.SharedString,
                    new CellValue(GetSharedString(instanceContext, reader)));
                break;
            case RecordTypeIdentifier.CELLERROR or RecordTypeIdentifier.CELLFMLAERROR:
                (cellType, cellValue) = (CellType.Error, new CellValue((ExcelErrorCode)reader.GetByte(8)));
                break;
        }

        // Return null for unhandled record types
        if (cellType == CellType.Unknown 
            || cellValue == null
            //|| recordType != RecordTypeIdentifier.CELLRK
            )
        {
            return null;
        }
        // Allocate once and populate properties
        return new XlsbCell
        {
            ExcelColumnOffset = columnOffset,
            RawExcelType = cellType,
            CellValue = cellValue.Value
        };
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static string? GetSharedString(InstanceContext instanceContext, PooledRecordBuffer reader) => instanceContext.SharedStrings?[reader.GetInt32(8)];

    [MethodImpl(MethodImplOptions.AggressiveInlining | MethodImplOptions.AggressiveOptimization)]
    private static double MagicConvertRK(PooledRecordBuffer record)
    {
        int rk = record.GetInt32(8);

        // Extract and process the RK value in a single optimized path
        bool isFloat = (rk & 0x02) == 0;
        double d;

        if (isFloat)
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

        return d;
    }

    /// <InheritDoc />
    public CellValue CellValue { get; private init; }

    /// <InheritDoc />
    public CellType RawExcelType { get; private init; }

    /// <InheritDoc />
    public IReadOnlyList<char> ColumnLetters
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get
        {
            int offset = ExcelColumnOffset;
            if (offset <= 0)
            {
                return Array.Empty<char>();
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
                char[] result = offset.GetExcelColumnName().ToCharArray();
                s_columnLetterCache[offset] = result;
                return result;
            }
            // Fallback for out-of-range offsets
            return offset.GetExcelColumnName().ToCharArray();
        }
    }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; private init; }
}