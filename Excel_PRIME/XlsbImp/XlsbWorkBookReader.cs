using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;

namespace ExcelPRIME.Implementation;

internal sealed class XlsbWorkBookReader : IOpenXmlWorkBookReaderAsync
{
    private readonly BufferedStream _stream;
    private readonly XlsbStreamReader _reader;
    private bool _isDisposed;

    public XlsbWorkBookReader(Stream stream, CancellationToken _)
    {
        // For modern hardware in 2025, 65536(64KB) is the standard "sweet spot" for many workloads
        _stream = new BufferedStream(stream!, 64 * 1024);
        _reader = new XlsbStreamReader(_stream);
    }

    public async IAsyncEnumerable<KeyValuePair<string, int>> GetSheetNamesAsync(
        [EnumeratorCancellation] CancellationToken ct)
    {
        int relativeSheetId = 0;

        PooledRecordBuffer nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        bool foundSheets = false;
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BUNDLESHEET)
            {
                foundSheets = true;
                string? rel = nextRecord.GetString(8, out int next);
                string name = nextRecord.GetString(next);
                if (rel == null)
                {
                    // no sheet rel means it is a macro.
                }
                else
                {
                    relativeSheetId++;
                    // `r:id` and `sheetId` are not to be trusted
                    yield return new KeyValuePair<string, int>(name, relativeSheetId);
                }
            }

            nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
            if (foundSheets
                && nextRecord.RecordType != RecordTypeIdentifier.BUNDLESHEET
                )
            {
                break;
            }
        }
        nextRecord.Dispose();
    }

    public IEnumerable<KeyValuePair<string, int>> GetSheetNames([EnumeratorCancellation] CancellationToken ct)
    {
        int relativeSheetId = 0;

        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        bool foundSheets = false;
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BUNDLESHEET)
            {
                foundSheets = true;
                string? rel = nextRecord.GetString(8, out int next);
                string name = nextRecord.GetString(next);
                if (rel == null)
                {
                    // no sheet rel means it is a macro.
                }
                else
                {
                    relativeSheetId++;
                    // `r:id` and `sheetId` are not to be trusted
                    yield return new KeyValuePair<string, int>(name, relativeSheetId);
                }
            }

            nextRecord = _reader.ReadNextRecord();
            if (foundSheets
                && nextRecord.RecordType != RecordTypeIdentifier.BUNDLESHEET
               )
            {
                break;
            }
        }
        nextRecord.Dispose();
    }

    public async Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(
    IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        Dictionary<string, DefinedRange> definedRanges = [];
        List<string>? sheetRefs = null;

        PooledRecordBuffer nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        bool foundDefinedNames = false;
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BRTNAME)
            {
                foundDefinedNames = true;
                int localSheetId = nextRecord.GetInt32(5);
                string name = nextRecord.GetString(9, out int formulaBegin)!;


                (string columnStart, string columnEnd, int rowStart, int rowEnd, bool isNumber, short sheetRef) = DecodeNameParsedFormula(nextRecord, formulaBegin);
                if (sheetRef < 0
                    && localSheetId > -1)
                {
                    sheetRef = (short)localSheetId;
                }
                string keyName = name;
                string sheetNameRef = string.Empty;
                if (sheetRef > -1)
                {
                    sheetRefs ??= [.. sheetNamesToOffsetSheetId.Keys];
                    sheetNameRef = sheetRefs[sheetRef];
                    if (!string.IsNullOrEmpty(sheetNameRef))
                    {
                        keyName = string.Concat(name, " (", sheetNameRef, ")");
                    }
                }

                definedRanges[keyName] = isNumber
                    ? new DefinedRange(columnStart) { Name = name, SheetIdReference = sheetNameRef }
                    : new DefinedRange(sheetNameRef, columnStart, columnEnd, rowStart, rowEnd) { Name = name, SheetIdReference = sheetNameRef };
                if (localSheetId < 0)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }

            nextRecord = await _reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
            if (foundDefinedNames
                && nextRecord.RecordType != RecordTypeIdentifier.BRTNAME)
            {
                break;
            }
        }
        nextRecord.Dispose();

        definedRanges.TrimExcess();
        return new ReadOnlyDictionary<string, DefinedRange>(definedRanges);
    }

    public IReadOnlyDictionary<string, DefinedRange> GetDefinedRanges(
        IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        Dictionary<string, DefinedRange> definedRanges = [];
        List<string>? sheetRefs = null;

        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        bool foundDefinedNames = false;
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BRTNAME)
            {
                foundDefinedNames = true;
                int localSheetId = nextRecord.GetInt32(5);
                string name = nextRecord.GetString(9, out int formulaBegin)!;

                (string columnStart, string columnEnd, int rowStart, int rowEnd, bool isNumber, short sheetRef) = DecodeNameParsedFormula(nextRecord, formulaBegin);
                if (sheetRef >= 0
                    && localSheetId == -1)
                {
                    localSheetId = sheetRef;
                }
                string keyName = name;
                string sheetNameRef = string.Empty;
                if (localSheetId != -1)
                {
                    sheetRefs ??= [.. sheetNamesToOffsetSheetId.Keys];
                    sheetNameRef = sheetRefs[localSheetId];
                    if (!string.IsNullOrEmpty(sheetNameRef))
                    {
                        keyName = string.Concat(name, " (", sheetNameRef, ")");
                    }
                }

                definedRanges[keyName] = isNumber
                    ? new DefinedRange(columnStart) { Name = name, SheetIdReference = sheetNameRef }
                    : new DefinedRange(sheetNameRef, columnStart, columnEnd, rowStart, rowEnd) { Name = name, SheetIdReference = sheetNameRef };
                if (definedRanges[keyName].SheetName == sheetNameRef)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }

            nextRecord = _reader.ReadNextRecord();
            if (foundDefinedNames
                && nextRecord.RecordType != RecordTypeIdentifier.BRTNAME)
            {
                break;
            }
        }
        nextRecord.Dispose();

        definedRanges.TrimExcess();
        return new ReadOnlyDictionary<string, DefinedRange>(definedRanges);
    }

    private static (string columnStart, string columnEnd, int rowStart, int rowEnd, bool isNumber, short sheetRef) DecodeNameParsedFormula(PooledRecordBuffer nextRecord, int formulaBegin)
    {
        int cce = nextRecord.GetInt32(formulaBegin);
        // PtgRef -> 0x24
        // PtgArea -> 0x25
        // PtgRefN -> 0x2C
        // PtgAreaN -> 0x2D
        // PtgRef -> 0x44
        // PtgArea -> 0x45
        // PtgRef -> 0x64
        // PtgArea -> 0x65

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
                    return (col.GetExcelColumnName(), col.GetExcelColumnName(), row, row, false, sheetRef);
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
        //offset = formulaBegin + cce;
        //int cb = nextRecord.GetInt32(offset);
        //if (cb > 0)
        //{
        //    sheetRef = nextRecord.GetInt16(offset + 5);
        //}

        return (colFirst.GetExcelColumnName(), colLast.GetExcelColumnName(), rowFirst, rowLast, false, sheetRef);
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _stream?.Dispose();
            }

            _isDisposed = true;
        }
    }

    ~XlsbWorkBookReader()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }
}
