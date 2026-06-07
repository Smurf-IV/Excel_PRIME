using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.XlsbImp;


namespace ExcelPRIME.Implementation;

internal class XlsbWorkBookReader : IOpenXmlWorkBookReader
{
    private protected readonly IZipReader _zipReader;
    private protected BufferedStream _streamWb;
    private protected XlsbStreamReader _readerWb;
    private bool _isDisposed;

    protected XlsbWorkBookReader(IZipReader zipReader)
    {
        ArgumentNullException.ThrowIfNull(zipReader);
        _zipReader = zipReader;
    }

    public XlsbWorkBookReader(IZipReader zipReader, CancellationToken _)
        : this(zipReader)
    {
        Stream? stream = zipReader.GetEntry("xl/workbook.bin");
        // For modern hardware in 2025, 65536(64KB) is the standard "sweet spot" for many workloads
        _streamWb = new BufferedStream(stream!, 64 * 1024);
        OpenWorkbookStream();
    }

    protected void OpenWorkbookStream() =>
        _readerWb = new XlsbStreamReader(_streamWb);

    public IEnumerable<KeyValuePair<string, string>> GetSheetNames([EnumeratorCancellation] CancellationToken ct)
    {
        Dictionary<string, string> worksheetRels = PopulateWorkSheetRels(ct);
        if (!worksheetRels.Any())
        {
            yield break;
        }

        PooledRecordBuffer nextRecord = _readerWb.ReadNextRecord();
        const string xl = "xl/";
        bool foundSheets = false;
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BUNDLESHEET)
            {
                foundSheets = true;
                string? rel = nextRecord.GetString(8, out int next);
                string sheetName = nextRecord.GetString(next);
                if (rel == null)
                {
                    // no sheet rel means it is a macro.
                }
                else
                {
                    yield return new KeyValuePair<string, string>(sheetName, xl + worksheetRels[rel]);
                }
            }

            nextRecord = _readerWb.ReadNextRecord();
            if (foundSheets
                && nextRecord.RecordType != RecordTypeIdentifier.BUNDLESHEET
               )
            {
                break;
            }
        }
        nextRecord.Dispose();
    }

    private Dictionary<string, string> PopulateWorkSheetRels(CancellationToken ct)
    {
        using Stream streamRelWb = _zipReader.GetEntry("xl/_rels/workbook.bin.rels")!;
        using XmlReader readerRels = XmlReader.Create(streamRelWb, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, // Disable DTDs for untrusted sources
            IgnoreComments = true, // Skip parsing and allocating strings for comments
            IgnoreWhitespace = true, // Ignore significant whitespace
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            //NameTable = new WorkBookRelsRestrictedNameTable(),
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });
        Dictionary<string, string> worksheetRels = [];
        string relationshipsRefAtom = readerRels.NameTable.Add("Relationships");
        if (!readerRels.ReadToFollowing(relationshipsRefAtom))
        {
            return worksheetRels;
        }
        string relationshipRefAtom = readerRels.NameTable.Add("Relationship");
        string idRefAtom = readerRels.NameTable.Add("Id");
        string targetRefAtom = readerRels.NameTable.Add("Target");
        while (readerRels.Read()
               && !readerRels.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (readerRels.NodeType == XmlNodeType.Element
                && readerRels.LocalName == relationshipRefAtom)
            {
                string id = string.Empty;
                string target = string.Empty;
                int expectedAttributes = 2;

                while (readerRels.MoveToNextAttribute() && expectedAttributes > 0)
                {
                    // Retrieve the atomized name directly.
                    string currentAttributeName = readerRels.LocalName;
                    if (ReferenceEquals(currentAttributeName, idRefAtom))
                    {
                        id = readerRels.Value;
                        expectedAttributes--;
                    }
                    else if (ReferenceEquals(currentAttributeName, targetRefAtom))
                    {
                        target = readerRels.Value;
                        expectedAttributes--;
                    }
                }
                if (expectedAttributes == 0)
                {
                    worksheetRels[id] = target;
                }
            }
        }
        return worksheetRels;
    }

    public IReadOnlyDictionary<string, DefinedRange> GetDefinedRanges(
        IReadOnlyDictionary<string, string> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        Dictionary<string, DefinedRange> definedRanges = [];
        List<string>? sheetRefs = null;

        PooledRecordBuffer nextRecord = _readerWb.ReadNextRecord();
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
                    ? new DefinedRange(columnStart) { Name = name }
                    : new DefinedRange(sheetNameRef, columnStart, columnEnd, rowStart, rowEnd) { Name = name};
                if (definedRanges[keyName].SheetName == sheetNameRef)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }

            nextRecord = _readerWb.ReadNextRecord();
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
                    string colName = new(col.GetExcelColumnName());
                    return (colName, colName, row, row, false, sheetRef);
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

        return (new string(colFirst.GetExcelColumnName()), new string(colLast.GetExcelColumnName()), rowFirst, rowLast, false, sheetRef);
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _streamWb.Dispose();
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

internal sealed class XlsbWorkBookReaderAsync : XlsbWorkBookReader, IOpenXmlWorkBookReaderAsync
{
    // ReSharper disable InconsistentNaming
    private IZipReaderAsync _zipReaderA => (IZipReaderAsync)base._zipReader;
    // ReSharper restore InconsistentNaming

    internal XlsbWorkBookReaderAsync(IZipReaderAsync zipReader)
#pragma warning disable CA2016 // do not forward the ct to the public base constructor
        : base(zipReader)
#pragma warning restore CA2016
    {
    }

    internal async Task InitializeAsync(CancellationToken ct)
    {
        Stream? stream = await _zipReaderA.GetEntryAsync("xl/workbook.bin", ct).ConfigureAwait(false);
        // For modern hardware in 2025, 65536(64KB) is the standard "sweet spot" for many workloads
        _streamWb = new BufferedStream(stream!, 64 * 1024);
        OpenWorkbookStream();
    }

    public async IAsyncEnumerable<KeyValuePair<string, string>> GetSheetNamesAsync([EnumeratorCancellation] CancellationToken ct)
    {
        Dictionary<string, string> worksheetRels = await PopulateWorkSheetRelsAsync(ct).ConfigureAwait(false);
        if (!worksheetRels.Any())
        {
            yield break;
        }

        PooledRecordBuffer nextRecord = await _readerWb.ReadNextRecordAsync(ct).ConfigureAwait(false);
        bool foundSheets = false;
        const string xl = "xl/";
        while (nextRecord.Succeeded
               && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.BUNDLESHEET)
            {
                foundSheets = true;
                string? rel = nextRecord.GetString(8, out int next);
                string sheetName = nextRecord.GetString(next);
                if (rel == null)
                {
                    // no sheet rel means it is a macro.
                }
                else
                {
                    yield return new KeyValuePair<string, string>(sheetName, xl + worksheetRels[rel]);
                }
            }

            nextRecord = await _readerWb.ReadNextRecordAsync(ct).ConfigureAwait(false);
            if (foundSheets
                && nextRecord.RecordType != RecordTypeIdentifier.BUNDLESHEET
               )
            {
                break;
            }
        }
        nextRecord.Dispose();
    }

    private async Task<Dictionary<string, string>> PopulateWorkSheetRelsAsync(CancellationToken ct)
    {
        using Stream? streamRelWb = await _zipReaderA.GetEntryAsync("xl/_rels/workbook.bin.rels", ct).ConfigureAwait(false);
        using XmlReader readerRels = XmlReader.Create(streamRelWb!, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, // Disable DTDs for untrusted sources
            IgnoreComments = true, // Skip parsing and allocating strings for comments
            IgnoreWhitespace = true, // Ignore significant whitespace
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            //NameTable = new WorkBookRelsRestrictedNameTable(),
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });
        Dictionary<string, string> worksheetRels = [];
        string relationshipsRefAtom = readerRels.NameTable.Add("Relationships");
        if (!await readerRels.ReadToFollowingAsync(relationshipsRefAtom).ConfigureAwait(false))
        {
            return worksheetRels;
        }
        string relationshipRefAtom = readerRels.NameTable.Add("Relationship");
        string idRefAtom = readerRels.NameTable.Add("Id");
        string targetRefAtom = readerRels.NameTable.Add("Target");
        while (await readerRels.ReadAsync().ConfigureAwait(false)
               && !readerRels.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (readerRels.NodeType == XmlNodeType.Element
                && readerRels.LocalName == relationshipRefAtom)
            {
                string id = string.Empty;
                string target = string.Empty;
                int expectedAttributes = 2;

                while (readerRels.MoveToNextAttribute() && expectedAttributes > 0)
                {
                    // Retrieve the atomized name directly.
                    string currentAttributeName = readerRels.LocalName;
                    if (ReferenceEquals(currentAttributeName, idRefAtom))
                    {
                        id = readerRels.Value;
                        expectedAttributes--;
                    }
                    else if (ReferenceEquals(currentAttributeName, targetRefAtom))
                    {
                        target = readerRels.Value;
                        expectedAttributes--;
                    }
                }
                if (expectedAttributes == 0)
                {
                    worksheetRels[id] = target;
                }
            }
        }
        return worksheetRels;
    }

    public async Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(
        IReadOnlyDictionary<string, string> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        Dictionary<string, DefinedRange> definedRanges = [];
        List<string>? sheetRefs = null;

        PooledRecordBuffer nextRecord = await _readerWb.ReadNextRecordAsync(ct).ConfigureAwait(false);
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
                    ? new DefinedRange(columnStart) { Name = name }
                    : new DefinedRange(sheetNameRef, columnStart, columnEnd, rowStart, rowEnd) { Name = name };
                if (localSheetId < 0)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }

            nextRecord = await _readerWb.ReadNextRecordAsync(ct).ConfigureAwait(false);
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
        //offset = formulaBegin + cce;
        //int cb = nextRecord.GetInt32(offset);
        //if (cb > 0)
        //{
        //    sheetRef = nextRecord.GetInt16(offset + 5);
        //}

        string columnName = new(colLast.GetExcelColumnName());
        return (new string(colFirst.GetExcelColumnName()), columnName, rowFirst, rowLast, false, sheetRef);
    }
}
