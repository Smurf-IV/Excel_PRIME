using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;


namespace ExcelPRIME.Implementation;

internal class XmlWorkBookReader : IOpenXmlWorkBookReader
{
    private protected readonly IZipReader _zipReader;
    private protected Stream _streamWb;
    private protected XmlReader _readerWb;
    private bool _isDisposed;

    protected XmlWorkBookReader(IZipReader zipReader) => _zipReader = zipReader;

    public XmlWorkBookReader(IZipReader zipReader, CancellationToken _)
    : this(zipReader)
    {
        _streamWb = _zipReader.GetEntry("xl/workbook.xml")!;
        OpenWorkbookStream();
    }

    protected void OpenWorkbookStream() =>
        _readerWb = XmlReader.Create(_streamWb, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, // Disable DTDs for untrusted sources
            IgnoreComments = true, // Skip parsing and allocating strings for comments
            IgnoreWhitespace = true, // Ignore significant whitespace
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            NameTable = new WorkBookRestrictedNameTable(),
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });

    public IEnumerable<KeyValuePair<string, string>> GetSheetNames([EnumeratorCancellation] CancellationToken ct)
    {
        Dictionary<string, string> worksheetRels = PopulateWorkSheetRels(ct);

        string sheetsRefAtom = _readerWb.NameTable.Add("sheets");
        if (!worksheetRels.Any()
            || !_readerWb.ReadToFollowing(sheetsRefAtom))
        {
            yield break;
        }

        string nameRefAtom = _readerWb.NameTable.Add("name");
        string sheetRefAtom = _readerWb.NameTable.Add("sheet");
        string idRefAtom = _readerWb.NameTable.Add("r:id");
        const string xl = "xl/";

        while (_readerWb.Read()
               && !_readerWb.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_readerWb.NodeType == XmlNodeType.Element
                && _readerWb.LocalName == sheetRefAtom)
            {
                if (_readerWb.MoveToAttribute(nameRefAtom))
                {
                    string sheetName = _readerWb.Value;
                    if (_readerWb.MoveToAttribute(idRefAtom))
                    {
                        yield return new KeyValuePair<string, string>(sheetName, xl + worksheetRels[_readerWb.Value]);
                    }
                }
            }
        }
    }

    private Dictionary<string, string> PopulateWorkSheetRels(CancellationToken ct)
    {
        using Stream streamRelWb = _zipReader.GetEntry("xl/_rels/workbook.xml.rels")!;
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
        string definedNamesRefAtom = _readerWb.NameTable.Add("definedNames");
        Dictionary<string, DefinedRange> definedRanges = [];
        if (!_readerWb.ReadToFollowing(definedNamesRefAtom))
        {
            definedRanges.TrimExcess();
            return definedRanges;
        }

        string definedNameRefAtom = _readerWb.NameTable.Add("definedName");
        string nameRefAtom = _readerWb.NameTable.Add("name");
        string localSheetIdRefAtom = _readerWb.NameTable.Add("localSheetId");
        List<string>? sheetRefs = null;
        while (_readerWb.Read()
               && !_readerWb.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_readerWb.NodeType == XmlNodeType.Element
                && _readerWb.LocalName == definedNameRefAtom)
            {
                string name = string.Empty;
                string localSheetId = string.Empty;
                int expectedAttributes = 2;

                while (_readerWb.MoveToNextAttribute() && expectedAttributes > 0)
                {
                    // Retrieve the atomized name directly.
                    string currentAttributeName = _readerWb.LocalName;
                    if (ReferenceEquals(currentAttributeName, nameRefAtom))
                    {
                        name = _readerWb.Value;
                        expectedAttributes--;
                    }
                    else if (ReferenceEquals(currentAttributeName, localSheetIdRefAtom))
                    {
                        localSheetId = _readerWb.Value;
                        expectedAttributes--;
                    }
                }

                string keyName = name;
                string sheetRef = string.Empty;
                if (!string.IsNullOrWhiteSpace(localSheetId))
                {
                    sheetRefs ??= [.. sheetNamesToOffsetSheetId.Keys];
                    int sheetId = localSheetId.IntParse();
                    sheetRef = sheetRefs[sheetId];
                    if (!string.IsNullOrEmpty(sheetRef))
                    {
                        keyName = string.Concat(name, " (", sheetRef, ")");
                    }
                }

                // Move to data
                _readerWb.Read();
                definedRanges[keyName] = new DefinedRange(_readerWb.Value) { Name = name };
                // Handle this situation-> <definedName name="DışVeri_2" localSheetId="3" hidden="1">Worksheet!$A$952351:$H$985351</definedName>
                if (definedRanges[keyName].SheetName == sheetRef)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }
        }

        definedRanges.TrimExcess();
        return definedRanges;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _readerWb.Dispose();
                _streamWb.Dispose();
            }

            _isDisposed = true;
        }
    }

    ~XmlWorkBookReader()
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

internal sealed class XmlWorkBookReaderAsync : XmlWorkBookReader, IOpenXmlWorkBookReaderAsync
{
    // ReSharper disable InconsistentNaming
    private IZipReaderAsync _zipReaderA => (IZipReaderAsync)base._zipReader;
    // ReSharper restore InconsistentNaming

    public XmlWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct)
#pragma warning disable CA2016 // do not forward the ct to the public base constructor
        : base(zipReader)
#pragma warning restore CA2016
    {
        _streamWb = zipReader.GetEntryAsync("xl/workbook.xml", ct).GetAwaiter().GetResult()!;
        OpenWorkbookStream();
    }

    public async IAsyncEnumerable<KeyValuePair<string, string>> GetSheetNamesAsync([EnumeratorCancellation] CancellationToken ct)
    {
        Dictionary<string, string> worksheetRels = await PopulateWorkSheetRelsAsync(ct).ConfigureAwait(false);

        string sheetsRefAtom = _readerWb.NameTable.Add("sheets");
        if (!worksheetRels.Any()
            || !_readerWb.ReadToFollowing(sheetsRefAtom))
        {
            yield break;
        }

        string nameRefAtom = _readerWb.NameTable.Add("name");
        string sheetRefAtom = _readerWb.NameTable.Add("sheet");
        string idRefAtom = _readerWb.NameTable.Add("r:id");
        const string xl = "xl/";
        while (await _readerWb.ReadAsync().ConfigureAwait(false)
               && !_readerWb.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_readerWb.NodeType == XmlNodeType.Element
                && _readerWb.LocalName == sheetRefAtom)
            {
                if (_readerWb.MoveToAttribute(nameRefAtom))
                {
                    string sheetName = _readerWb.Value;
                    if (_readerWb.MoveToAttribute(idRefAtom))
                    {
                        yield return new KeyValuePair<string, string>(sheetName, xl + worksheetRels[_readerWb.Value]);
                    }
                }
            }
        }
    }

    private async Task<Dictionary<string, string>> PopulateWorkSheetRelsAsync(CancellationToken ct)
    {
        using Stream streamRelWb = await _zipReaderA.GetEntryAsync("xl/_rels/workbook.xml.rels", ct).ConfigureAwait(false)!;
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
        string definedNamesRefAtom = _readerWb.NameTable.Add("definedNames");
        Dictionary<string, DefinedRange> definedRanges = [];
        if (!_readerWb.ReadToFollowing(definedNamesRefAtom))
        {
            definedRanges.TrimExcess();
            return definedRanges;
        }

        string definedNameRefAtom = _readerWb.NameTable.Add("definedName");
        string nameRefAtom = _readerWb.NameTable.Add("name");
        string localSheetIdRefAtom = _readerWb.NameTable.Add("localSheetId");
        List<string>? sheetRefs = null;
        while (await _readerWb.ReadAsync().ConfigureAwait(false)
               && !_readerWb.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_readerWb.NodeType == XmlNodeType.Element
                && _readerWb.LocalName == definedNameRefAtom)
            {
                string name = string.Empty;
                string localSheetId = string.Empty;
                int expectedAttributes = 2;

                while (_readerWb.MoveToNextAttribute() && expectedAttributes > 0)
                {
                    // Retrieve the atomized name directly.
                    string currentAttributeName = _readerWb.LocalName;
                    if (ReferenceEquals(currentAttributeName, nameRefAtom))
                    {
                        name = _readerWb.Value;
                        expectedAttributes--;
                    }
                    else if (ReferenceEquals(currentAttributeName, localSheetIdRefAtom))
                    {
                        localSheetId = _readerWb.Value;
                        expectedAttributes--;
                    }
                }

                string keyName = name;
                string sheetRef = string.Empty;
                if (!string.IsNullOrWhiteSpace(localSheetId))
                {
                    sheetRefs ??= [.. sheetNamesToOffsetSheetId.Keys];
                    int sheetId = localSheetId.IntParse();
                    sheetRef = sheetRefs[sheetId];
                    if (!string.IsNullOrEmpty(sheetRef))
                    {
                        keyName = string.Concat(name, " (", sheetRef, ")");
                    }
                }

                // Move to data
                await _readerWb.ReadAsync().ConfigureAwait(false);
                definedRanges[keyName] = new DefinedRange(_readerWb.Value) { Name = name };
                // Handle this situation-> <definedName name="DışVeri_2" localSheetId="3" hidden="1">Worksheet!$A$952351:$H$985351</definedName>
                if (definedRanges[keyName].SheetName == sheetRef)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }
        }

        definedRanges.TrimExcess();
        return definedRanges;
    }

}
