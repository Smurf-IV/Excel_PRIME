using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Implementation;

internal sealed class XmlWorkBookReader : IOpenXmlWorkBookReaderAsync
{
    private readonly Stream _stream;
    private readonly XmlReader _reader;
    private bool _isDisposed;

    public XmlWorkBookReader(Stream? stream, CancellationToken _)
    {
        _stream = stream!;
        _reader = XmlReader.Create(_stream, new XmlReaderSettings
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
    }

    public async IAsyncEnumerable<KeyValuePair<string, int>> GetSheetNamesAsync([EnumeratorCancellation] CancellationToken ct)
    {
        string sheetsRefAtom = _reader.NameTable.Add("sheets");
        if (!_reader.ReadToFollowing(sheetsRefAtom))
        {
            yield break;
        }

        string nameRefAtom = _reader.NameTable.Add("name");
        string sheetRefAtom = _reader.NameTable.Add("sheet");
        int relativeSheetId = 0;
        while (await _reader.ReadAsync().ConfigureAwait(false)
               && !_reader.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && _reader.LocalName == sheetRefAtom)
            {
                if (_reader.MoveToAttribute(nameRefAtom))
                {
                    relativeSheetId++;
                    // `r:id` and `sheetId` are not to be trusted
                    yield return new KeyValuePair<string, int>(_reader.Value, relativeSheetId);
                }
            }
        }
    }

    public IEnumerable<KeyValuePair<string, int>> GetSheetNames([EnumeratorCancellation] CancellationToken ct)
    {
        string sheetsRefAtom = _reader.NameTable.Add("sheets");
        if (!_reader.ReadToFollowing(sheetsRefAtom))
        {
            yield break;
        }

        string nameRefAtom = _reader.NameTable.Add("name");
        string sheetRefAtom = _reader.NameTable.Add("sheet");
        int relativeSheetId = 0;
        while (_reader.Read()
               && !_reader.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && _reader.LocalName == sheetRefAtom)
            {
                if (_reader.MoveToAttribute(nameRefAtom))
                {
                    relativeSheetId++;
                    // `r:id` and `sheetId` are not to be trusted
                    yield return new KeyValuePair<string, int>(_reader.Value, relativeSheetId);
                }
            }
        }
    }

    public async Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(
        IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        string definedNamesRefAtom = _reader.NameTable.Add("definedNames");
        Dictionary<string, DefinedRange> definedRanges = [];
        if (!_reader.ReadToFollowing(definedNamesRefAtom))
        {
            definedRanges.TrimExcess();
            return definedRanges.AsReadOnly();
        }

        string definedNameRefAtom = _reader.NameTable.Add("definedName");
        string nameRefAtom = _reader.NameTable.Add("name");
        string localSheetIdRefAtom = _reader.NameTable.Add("localSheetId");
        List<string>? sheetRefs = null;
        while (await _reader.ReadAsync().ConfigureAwait(false)
               && !_reader.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && _reader.LocalName == definedNameRefAtom)
            {
                string name = string.Empty;
                string localSheetId = string.Empty;
                int expectedAttributes = 2;

                while (_reader.MoveToNextAttribute() && expectedAttributes > 0)
                {
                    // Retrieve the atomized name directly.
                    string currentAttributeName = _reader.LocalName;
                    if (ReferenceEquals(currentAttributeName, nameRefAtom))
                    {
                        name = _reader.Value;
                        expectedAttributes--;
                    }
                    else if (ReferenceEquals(currentAttributeName, localSheetIdRefAtom))
                    {
                        localSheetId = _reader.Value;
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
                await _reader.ReadAsync().ConfigureAwait(false);
                definedRanges[keyName] = new DefinedRange(_reader.Value) { Name = name, SheetIdReference = sheetRef };
                // Handle this situation-> <definedName name="DışVeri_2" localSheetId="3" hidden="1">Worksheet!$A$952351:$H$985351</definedName>
                if (definedRanges[keyName].SheetName == sheetRef)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }
        }

        definedRanges.TrimExcess();
        return definedRanges.AsReadOnly();
    }

    public IReadOnlyDictionary<string, DefinedRange> GetDefinedRanges(IReadOnlyDictionary<string, int> sheetNamesToOffsetSheetId, CancellationToken ct)
    {
        string definedNamesRefAtom = _reader.NameTable.Add("definedNames");
        Dictionary<string, DefinedRange> definedRanges = [];
        if (!_reader.ReadToFollowing(definedNamesRefAtom))
        {
            definedRanges.TrimExcess();
            return definedRanges.AsReadOnly();
        }

        string definedNameRefAtom = _reader.NameTable.Add("definedName");
        string nameRefAtom = _reader.NameTable.Add("name");
        string localSheetIdRefAtom = _reader.NameTable.Add("localSheetId");
        List<string>? sheetRefs = null;
        while ( _reader.Read()
               && !_reader.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && _reader.LocalName == definedNameRefAtom)
            {
                string name = string.Empty;
                string localSheetId = string.Empty;
                int expectedAttributes = 2;

                while (_reader.MoveToNextAttribute() && expectedAttributes > 0)
                {
                    // Retrieve the atomized name directly.
                    string currentAttributeName = _reader.LocalName;
                    if (ReferenceEquals(currentAttributeName, nameRefAtom))
                    {
                        name = _reader.Value;
                        expectedAttributes--;
                    }
                    else if (ReferenceEquals(currentAttributeName, localSheetIdRefAtom))
                    {
                        localSheetId = _reader.Value;
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
                _reader.Read();
                definedRanges[keyName] = new DefinedRange(_reader.Value) { Name = name, SheetIdReference = sheetRef };
                // Handle this situation-> <definedName name="DışVeri_2" localSheetId="3" hidden="1">Worksheet!$A$952351:$H$985351</definedName>
                if (definedRanges[keyName].SheetName == sheetRef)
                {
                    definedRanges.TryAdd(name, definedRanges[keyName]);
                }
            }
        }

        definedRanges.TrimExcess();
        return definedRanges.AsReadOnly();
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _reader.Dispose();
                _stream?.Dispose();
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
