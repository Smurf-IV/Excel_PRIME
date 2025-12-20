using System;
using System.Collections.Generic;
using System.Diagnostics;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

[DebuggerDisplay("{RawValue.ToString(),raw}")]
internal sealed record XlsbCell : ICell
{
    public static XlsbCell? ConstructCell(PooledRecordBuffer reader, InstanceContext instanceContext)
    {
        XlsbCell cell = new XlsbCell
        {
            ExcelColumnOffset = reader.GetInt32(0) + 1 // Convert zero-based to one-based
        };

        switch (reader.RecordType)
        {
            //case RecordTypeIdentifier.CELLBLANK:
            //    cell.RawExcelType = CellType.Unknown;
            //    break;
            case RecordTypeIdentifier.CELLRK:
                cell.RawExcelType = CellType.Numeric;
                cell.RawValue = MagicConvertRK(reader);
                break;
            case RecordTypeIdentifier.CELLFMLAERROR:
            case RecordTypeIdentifier.CELLERROR:
                cell.RawExcelType = CellType.Error;
                // TODO: Decrypt the error
                cell.RawValue = reader.GetByte(8);
                break;
            case RecordTypeIdentifier.CELLFMLABOOL:
            case RecordTypeIdentifier.CELLBOOL:
                cell.RawExcelType = CellType.Boolean;
                cell.RawValue = reader.GetByte(8) != 0;
                break;
            case RecordTypeIdentifier.CELLFMLANUM:
            case RecordTypeIdentifier.CELLREAL:
                cell.RawExcelType = CellType.Numeric;
                cell.RawValue = reader.GetDouble(8);
                break;
            case RecordTypeIdentifier.CELLFMLASTRING:
            case RecordTypeIdentifier.CELLST:
                cell.RawExcelType = CellType.InlineString;
                cell.RawValue = reader.GetString(8);
                break;
            case RecordTypeIdentifier.CELLISST:
                cell.RawExcelType = CellType.SharedString;
                cell.RawValue = instanceContext.SharedStrings?[reader.GetInt32(8)];
                break;
            default:
                return null;
        }

        return cell;
    }

    private static double MagicConvertRK(PooledRecordBuffer record)
    {
        int rk = record.GetInt32(8);

        bool isFloat = (rk & 0x02) == 0;
        double d;

        if (isFloat)
        {
            long v = rk & 0xfffffffc;
            v <<= 32;
            d = BitConverter.Int64BitsToDouble(v);
        }
        else
        {
            d = rk >> 2;
        }

        bool isScaled = (rk & 0x01) != 0;
        if (isScaled)
        {
            d /= 100;
        }

        return d;
    }


    /// <InheritDoc />
    public object? RawValue { get; private set; }

    /// <InheritDoc />
    public CellType RawExcelType { get; private set; }

    /// <InheritDoc />
    public IReadOnlyList<char> ColumnLetters => ExcelColumnOffset.GetExcelColumnName().ToCharArray();

    /// <InheritDoc />
    public int ExcelColumnOffset { get; private init; }

}