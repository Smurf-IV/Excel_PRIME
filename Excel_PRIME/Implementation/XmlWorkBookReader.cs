using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelPRIME.Implementation;

internal sealed class XmlWorkBookReader : IXmlWorkBookReader
{
    private readonly Stream _stream;
    private readonly XmlReader _reader;
    private readonly string _relationshipNamespace;
    private bool _isDisposed;
    private readonly string _nameRefAtom;
    private readonly string _sheetRefAtom;

    public XmlWorkBookReader(Stream? stream, CancellationToken ct)
    {
        _stream = stream;
        _reader = XmlReader.Create(stream, new XmlReaderSettings
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
        // advance to the content
        string workbookRefAtom = _reader.NameTable.Add("workbook");
        _reader.ReadToFollowing(workbookRefAtom);
        //string xmlns_rRefAtom = _reader.NameTable.Add("xmlns:r");
        //_relationshipNamespace =
        //    _reader.GetAttribute(xmlns_rRefAtom) ?? string.Empty; /* == XmlNamespaces.RelationshipsOclc
        //        ? XmlNamespaces.RelationshipsOclc
        //        : XmlNamespaces.RelationshipsOpenXmlFormat;*/

        _nameRefAtom = _reader.NameTable.Add("name");
        _sheetRefAtom = _reader.NameTable.Add("sheet");
    }

    public async IAsyncEnumerable<KeyValuePair<string, int>> GetSheetNamesAsync(CancellationToken ct)
    {
        string sheetsRefAtom = _reader.NameTable.Add("sheets");
        if (!_reader.ReadToFollowing(sheetsRefAtom))
        {
            yield break;
        }

        int relativeSheetId = 0;
        while (await _reader.ReadAsync().ConfigureAwait(false)
               && !_reader.EOF
               && !ct.IsCancellationRequested
              )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && _reader.LocalName == _sheetRefAtom)
            {
                if (_reader.MoveToAttribute(_nameRefAtom))
                {
                    relativeSheetId++;
                    // `r:id` and `sheetId` are not to be trusted
                    yield return new KeyValuePair<string, int>(_reader.Value, relativeSheetId);
                }
            }
        }
    }

    public Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct)
    {
        throw new NotImplementedException();
        //var definedNames = _document.Descendants()
        //    .Where(d => d.Name.LocalName == "definedNames")
        //    .Descendants();
        //Dictionary<string, DefinedRange> dict = new();
        //foreach (XElement e in definedNames)
        //{
        //    int worksheetIndex = -1;
        //    if (!string.IsNullOrWhiteSpace(e.Attribute("localSheetId")?.Value))
        //    {
        //        try
        //        {
        //            worksheetIndex = e.Attribute("localSheetId").Value.IntParseUnsafe() + 1;
        //        }
        //        catch (Exception exception)
        //        {
        //            // In a well-formed file, this should never happen.
        //            throw new KeyNotFoundException(string.Concat("Error reading localSheetId value for DefinedName: ", e.Attribute("name")?.Value), exception);
        //        }
        //    }

        //    DefinedRange range = new DefinedRange
        //    {
        //        Name = e.Attribute("name")?.Value ?? string.Empty,
        //        Reference = e.Value,
        //        SheetIndex = worksheetIndex
        //    };

        //    dict.Add(range.Key, range);
        //}

        //IReadOnlyDictionary<string, DefinedRange> readOnlyDictionary = dict.AsReadOnly();
        //return Task.FromResult(readOnlyDictionary);
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
