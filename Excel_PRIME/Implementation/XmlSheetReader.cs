using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Implementation;

internal sealed class XmlSheetReader : IOpenXmlSheetReaderAsync
{
    private readonly InstanceContext _instanceContext;
    private readonly XmlReader _reader;
    private bool _isDisposed;
    private readonly int _startRow;
    private readonly string _rowRefAtom;
    private readonly ReaderAtoms _readerAtoms;

    public XmlSheetReader(NonClosingStream stream, InstanceContext instanceContext, XmlNameTable sharedNameTable, CancellationToken ct)
    {
        _instanceContext = instanceContext;
        _reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, // Disable DTDs for untrusted sources
            IgnoreComments = true, // Skip parsing and allocating strings for comments
            IgnoreWhitespace = true, // Ignore insignificant whitespace
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            NameTable = sharedNameTable,
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });
        string worksheetRefAtom = _reader.NameTable.Add("worksheet");
        // Step into the worksheet
        while (_reader.Read() && !ct.IsCancellationRequested)
        {
            if (_reader.NodeType == XmlNodeType.Element
                && ReferenceEquals(_reader.LocalName, worksheetRefAtom)
               )
            {
                break;
            }
        }

        string dimensionRefAtom = _reader.NameTable.Add("dimension");
        string colsRefAtom = _reader.NameTable.Add("cols");
        string sheetDataRefAtom = _reader.NameTable.Add("sheetData");

        bool foundSheetData = false;
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

            if (ReferenceEquals(readerLocalName, dimensionRefAtom))
            {
                string? dim = _reader.GetAttribute("ref");
                if (dim != null)
                {
                    ReadOnlySpan<char> dimSpan = dim.AsSpan();
                    int colonIndex = dimSpan.IndexOf(':');
                    if (colonIndex == -1)
                    {
                        (int rowExcel, int _, _) = dimSpan.GetRowColNumbers();
                        _startRow = rowExcel - 1; // Take it back to the array offset
                        SheetDimensions = (1, 1);
                    }
                    else
                    {
                        ReadOnlySpan<char> firstPart = dimSpan[..colonIndex];
                        ReadOnlySpan<char> secondPart = dimSpan[(colonIndex + 1)..];
                        (int rowExcel, int _, _) = firstPart.GetRowColNumbers();
                        _startRow = rowExcel - 1; // Take it back to the array offset
                        (int rowMax, int colMax, _) = secondPart.GetRowColNumbers();
                        SheetDimensions = (rowMax, colMax);
                    }
                }
                else
                {
                    SheetDimensions = new ValueTuple<int, int>(0, 0);
                }
            }
            else if (ReferenceEquals(readerLocalName, colsRefAtom))
            {
                if (_reader.IsEmptyElement)
                {
                    // TODO: Need to understand when and how this is used
                    //continue;
                }
            }
            else if (ReferenceEquals(readerLocalName, sheetDataRefAtom))
            {
                foundSheetData = true;
            }
        }
        CurrentRow = 0;
        // Atomize key names once for fast lookups later.
        _rowRefAtom = sharedNameTable.Add("row");
        _readerAtoms = new ReaderAtoms(_reader);
    }

    private static Row CreateRowFromPool()
        => Row.Rent();

    private bool ReadToNextStartRow(CancellationToken ct)
    {
        while (_reader.ReadToFollowing(_rowRefAtom)
               && !ct.IsCancellationRequested
              )
        {
            if (_reader.NodeType == XmlNodeType.Element
                && ReferenceEquals(_reader.LocalName, _rowRefAtom)
               )
            {
                return true;
            }
        }
        if (_reader.EOF)
        {   // No rows to read, or the Dimension is lying
        }
        return false;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _lastNullRow = null;    // Do not call dispose, because they have been returned to the caller
                _lastRow = null;    // Do not call dispose, because they have been returned to the caller
                _reader.Dispose();
            }

            _isDisposed = true;
        }
    }

    public (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// The Current row iterator offset (Starts at 1)
    /// </summary>
    public int CurrentRow { get; private set; }

    public void Dispose() => Dispose(isDisposing: true);

#pragma warning disable CA2213  // Do not call dispose, because they are being returned to the caller
    private NullRow? _lastNullRow;
    private Row? _lastRow;
#pragma warning restore CA2213

    public async Task<IRowAsync?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        CurrentRow++;
        if (_lastRow != null)
        {
            if (_lastRow.RowOffset > CurrentRow)
            {
                _lastNullRow = new NullRow(CurrentRow);
                return _lastNullRow;
            }
            else
            {
                _lastNullRow = null;    // Do not call dispose, because they are being returned to the caller
                Row thisRow = _lastRow;
                _lastRow = null;    // Do not call dispose, because they are being returned to the caller
                return thisRow;
            }
        }

        if (CurrentRow < _startRow
            || !ReadToNextStartRow(ct)
           )
        {
            return null;
        }

        Row nextRow = CreateRowFromPool();
        nextRow.Initialize(_reader, _instanceContext, SheetDimensions.Width, _readerAtoms);

        if (cellGetMode > RowCellGet.None)
        {
            await nextRow.GetCellsAsync(ct).ConfigureAwait(false);
        }

        if (nextRow.RowOffset > CurrentRow)
        {
            // Deal with blank rows in the sheet?
            // i.e. ones that do not have a definition in the xml! Therefore, will "Look like a jump"
            _lastRow = nextRow;
            _lastNullRow = new NullRow(CurrentRow);
            return _lastNullRow;
        }

        return nextRow;
    }

    public IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        CurrentRow++;
        if (_lastRow != null)
        {
            if (_lastRow.RowOffset > CurrentRow)
            {
                _lastNullRow = new NullRow(CurrentRow);
                return _lastNullRow;
            }
            else
            {
                _lastNullRow = null;    // Do not call dispose, because they are being returned to the caller
                Row thisRow = _lastRow;
                _lastRow = null;    // Do not call dispose, because they are being returned to the caller
                return thisRow;
            }
        }

        if (CurrentRow < _startRow
            || !ReadToNextStartRow(ct)
           )
        {
            return null;
        }

        Row nextRow = CreateRowFromPool();
        nextRow.Initialize(_reader, _instanceContext, SheetDimensions.Width, _readerAtoms);

        if (cellGetMode > RowCellGet.None)
        {
            nextRow.GetCells(ct);
        }

        if (nextRow.RowOffset > CurrentRow)
        {
            // Deal with blank rows in the sheet?
            // i.e. ones that do not have a definition in the xml! Therefore, will "Look like a jump"
            _lastRow = nextRow;
            _lastNullRow = new NullRow(CurrentRow);
            return _lastNullRow;
        }

        return nextRow;
    }
}