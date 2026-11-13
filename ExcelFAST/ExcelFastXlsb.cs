using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

using ExcelFAST.Implementation;
using ExcelFAST.XlsbImp;

using ExcelPRIME;
using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;
using ExcelPRIME.XlsbImp;

namespace ExcelFAST;

/// <summary>
/// Main entry point into the `Excel_PRIMEXlsb` API's
/// </summary>
public sealed class ExcelFastXlsb : ExcelFastBase, IExcelFast
{
    /// <InheritDoc />
    public ExcelFastXlsb(IOpenXmlReaderHelpers? xlsbReader = null, IZipReader? zipReader = null)
    : base(xlsbReader ?? new XlsbReaderHelpers(),
            zipReader ?? new ZipReader(),
            new RowFastXlsb()
            )
    {
    }

    /// <inheritdoc />
    protected override IRowFast CreateRowFast() => new RowFastXlsb();

    /// <inheritdoc />
    public override IReadOnlyDictionary<string, DefinedRange> LoadDefinedRanges(CancellationToken ct)
    {
        Stream? stream = _zipReader.GetEntry("xl/workbook.bin", ct);
        if (stream == null)
        {
            return new Dictionary<string, DefinedRange>(StringComparer.InvariantCultureIgnoreCase);
        }

        using BufferedStream streamWb = new(stream, 64 * 1024);
        XlsbStreamReader readerWb = new(streamWb);

        Dictionary<string, DefinedRange> definedRanges = new(StringComparer.InvariantCultureIgnoreCase);
        List<string>? sheetRefs = null;

        PooledRecordBuffer nextRecord = readerWb.ReadNextRecord();
        bool foundDefinedNames = false;
        while (nextRecord.Succeeded && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BRTNAME)
            {
                foundDefinedNames = true;
                int localSheetId = nextRecord.GetInt32(5);
                string name = nextRecord.GetString(9, out int formulaBegin)!;

                (string columnStart, string columnEnd, int rowStart, int rowEnd, bool isNumber, short sheetRef) = DecodeNameParsedFormula(nextRecord, formulaBegin);
                if (localSheetId > -1)
                {
                    sheetRef = (short)localSheetId;
                }
                string keyName = name;
                string sheetNameRef = string.Empty;
                if (sheetRef > -1)
                {
                    sheetRefs ??= [.. _sheetNamesToPathOffset.Keys];
                    sheetNameRef = sheetRef < sheetRefs.Count ? sheetRefs[sheetRef] : string.Empty;
                    if (!string.IsNullOrEmpty(sheetNameRef))
                    {
                        keyName = string.Concat(name, " (", sheetNameRef, ")");
                    }
                }

                DefinedRange dr = isNumber
                    ? new DefinedRange(columnStart) { Name = name }
                    : new DefinedRange(sheetNameRef, columnStart, columnEnd, rowStart, rowEnd) { Name = name };

                definedRanges[keyName] = dr;
                if (localSheetId < 0)
                {
                    definedRanges[name] = dr;
                }
                else
                {
                    _ = definedRanges.TryAdd(name, dr);
                }
            }

            nextRecord = readerWb.ReadNextRecord();
            if (foundDefinedNames && nextRecord.RecordType != RecordTypeIdentifier.BRTNAME)
            {
                break;
            }
        }
        nextRecord.Dispose();

        definedRanges.TrimExcess();
        return definedRanges;
    }

    private static (string columnStart, string columnEnd, int rowStart, int rowEnd, bool isNumber, short sheetRef) DecodeNameParsedFormula(PooledRecordBuffer nextRecord, int formulaBegin)
    {
        int offset = formulaBegin + 4;
        byte ptg = nextRecord.GetByte(offset);
        short sheetRef = -1;
        offset++; // Step over Ptg### preamble

        switch (ptg)
        {
            case 0x1D:  // ptgBool
                return ((nextRecord.GetByte(offset) != 0).ToString(CultureInfo.InvariantCulture), string.Empty, 0, 0, true, sheetRef);
            case 0x1E:  // ptgInt
                return (nextRecord.GetInt16(offset).ToString(CultureInfo.InvariantCulture), string.Empty, 0, 0, true, sheetRef);
            case 0x1F:  // ptgNum
                return (nextRecord.GetDouble(offset).ToString(CultureInfo.InvariantCulture), string.Empty, 0, 0, true, sheetRef);

            case 0x25: //PtgArea
            case 0x45:
            case 0x65:
                break;

            case 0x3A: // PtgRef3d
            case 0x5A:
            case 0x7A:
                {
                    sheetRef = nextRecord.GetInt16(offset);
                    offset += 2; // Step over PtgArea3d `SheetRef`
                    int row = nextRecord.GetInt32(offset) + 1;
                    offset += 4;
                    int col = nextRecord.GetInt16(offset) + 1;
                    string excelColumnName = new(col.GetExcelColumnName());
                    return (excelColumnName, excelColumnName, row, row, false, sheetRef);
                }

            case 0x3B:  //PtgArea3d
            case 0x5B:
            case 0x7B:
                sheetRef = nextRecord.GetInt16(offset);
                offset += 2; // Step over PtgArea3d `SheetRef`
                break;

            default: // 0x23(35) -> PtgName | 0x39(57) -> PtgNameX
                return (string.Empty, string.Empty, 0, 0, false, sheetRef);
        }
        int rowFirst = nextRecord.GetInt32(offset) + 1;
        offset += 4;
        int rowLast = nextRecord.GetInt32(offset) + 1;
        offset += 4;
        int colFirst = nextRecord.GetInt16(offset) + 1;
        offset += 2;
        int colLast = nextRecord.GetInt16(offset) + 1;
        return (new string(colFirst.GetExcelColumnName()), new string(colLast.GetExcelColumnName()), rowFirst, rowLast, false, sheetRef);
    }
}
