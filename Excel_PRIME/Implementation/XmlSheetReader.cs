using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class XmlSheetReader : IXmlSheetReader
{
    private readonly ISharedString? _sharedStrings;
    private readonly XmlNameTable _sharedNameTable;
    private readonly XmlReader _reader;
    private bool _isDisposed;
    private readonly int _startRow;
    private readonly string _rowName;

    public XmlSheetReader(Stream stream, ISharedString sharedStrings, XmlNameTable sharedNameTable, CancellationToken ct)
    {
        _sharedStrings = sharedStrings;
        _sharedNameTable = sharedNameTable;
        _reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, // Disable DTDs for untrusted sources
            IgnoreComments = true, // Skip parsing and allocating strings for comments
            IgnoreWhitespace = true, // Ignore significant whitespace
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            NameTable = sharedNameTable,
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });
        string worksheetRef = _reader.NameTable.Add("worksheet");
        // Step into the worksheet
        while (_reader.Read() && !ct.IsCancellationRequested)
        {
            if (_reader.NodeType == XmlNodeType.Element
                && Object.ReferenceEquals(_reader.LocalName, worksheetRef)
                )
            {
                break;
            }
        }

        string dimensionRef = _reader.NameTable.Add("dimension");
        string colsRef = _reader.NameTable.Add("cols");
        string sheetDataRef = _reader.NameTable.Add("sheetData");

        var foundSheetData = false;
        while (!ct.IsCancellationRequested
               && !foundSheetData   // Do not read after finding sheetData
               && _reader.Read()
              )
        {
            if (_reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            string readerLocalName = _reader.LocalName;

            if (Object.ReferenceEquals(readerLocalName, dimensionRef))
            {
                string? dim = _reader.GetAttribute("ref");
                if (dim != null)
                {
                    string[] idx = dim.Split(':');
                    (int rowExcel, int _, ReadOnlyMemory<char> _) = idx[0].GetRowColNumbers();
                    _startRow = rowExcel - 1; // Take it back to the array offset
                    // Might be an empty sheet (i.e. only "A1")
                    if (idx.Length == 1)
                    {
                        SheetDimensions = new ValueTuple<int, int>(1, 1);
                    }
                    else
                    {
                        (int rowMax, int colMax, ReadOnlyMemory<char> _) = idx[1].GetRowColNumbers();

                        SheetDimensions = new ValueTuple<int, int>(rowMax, colMax);
                    }
                }
                else
                {
                    SheetDimensions = new ValueTuple<int, int>(0, 0);
                }
            }
            else if (Object.ReferenceEquals(readerLocalName, colsRef))
            {
                if (_reader.IsEmptyElement)
                {
                    // TODO: Need to understand when and how this is used
                    continue;
                }
            }
            else if (Object.ReferenceEquals(readerLocalName, sheetDataRef))
            {
                foundSheetData = true;
            }
        }
        CurrentRow = 0;
        // Atomize key names once for fast lookups later.
        _rowName = _sharedNameTable.Add("row");
    }

    private async Task<bool> ReadToNextStartRowAsync(CancellationToken ct)
    {
        while (await _reader.ReadAsync().ConfigureAwait(false)
               && !ct.IsCancellationRequested
               )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && Object.ReferenceEquals(_reader.LocalName, _rowName)
                )
            {
                CurrentRow++;
                return true;
            }
        }
        if (_reader.EOF)
        {   // No rows to read, or the Dimension is lying
            CurrentRow++;
        }
        return false;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _reader.Dispose();
                _sharedStrings?.Dispose();
            }

            _isDisposed = true;
        }
    }

    public (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// The Current row iterator offset (Starts at 1)
    /// </summary>
    public int CurrentRow { get; private set; }

    ~XmlSheetReader()
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

    public async Task<IRow?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        if (CurrentRow < _startRow ||
            !await ReadToNextStartRowAsync(ct).ConfigureAwait(false)
            )
        {
            return null;
        }
        Row nextRow = new Row(_reader, _sharedStrings, SheetDimensions.Width);
        if (cellGetMode > RowCellGet.None)
        {
            await nextRow.GetCellsAsync(ct);
        }

        //if (nextRow.RowOffset > CurrentRow)
        //{
        //    // TODO: How to deal with blank rows in the sheet?
        //    // i.e. ones that do not have a definition in the xml! Therefore, will "Look like a jump"
        //    throw new IndexOutOfRangeException($"nextRow.RowOffset [{nextRow.RowOffset}] > CurrentRow [{CurrentRow}]");
        //}

        return nextRow;
    }

    public IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default) => GetNextRowAsync(cellGetMode, ct).GetAwaiter().GetResult();

    public Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct) => throw new NotImplementedException();
}
