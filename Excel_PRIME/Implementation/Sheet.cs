using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Diagnostics.CodeAnalysis;

using ExcelPRIME.FromExternal;


namespace ExcelPRIME.Implementation; 

internal sealed class Sheet : ISheetAsync
{
    private bool _isDisposed;
    [SuppressMessage("Microsoft.Usage", "CA2213:DisposableFieldsShouldBeDisposed", Justification = "IOpenXmlReaderHelpersAsync is owned by Excel_PRIME and shared across sheets; disposing it here would break shared resources.")]
    private readonly IOpenXmlReaderHelpersAsync _xmlSharedReaderHelper;
    private readonly InstanceContext _instanceContext;
    private readonly XmlNameTable _sharedNameTable;
    private readonly NonClosingStream _stream;
    private IOpenXmlSheetReaderAsync? _sheetReader;

    internal Sheet(Stream stream, IOpenXmlReaderHelpersAsync xmlSharedReaderHelper, string name, InstanceContext instanceContext)
    {
        _stream = new NonClosingStream(stream);
        _xmlSharedReaderHelper = xmlSharedReaderHelper;
        _instanceContext = instanceContext;
        _sharedNameTable = new SheetRestrictedNameTable();
        Name = name;
    }

    /// <InheritDoc />
    public string Name { get; }

    /// <inheritdoc/>
    public (int Height, int Width) SheetDimensions
    {
        get
        {
            _sheetReader ??= (IOpenXmlSheetReaderAsync)_xmlSharedReaderHelper.CreateSheetReader(_stream, _instanceContext, _sharedNameTable, CancellationToken.None);

            return _sheetReader.SheetDimensions;
        }
    }

    /// <inheritdoc/>
    public int CurrentRow => _sheetReader?.CurrentRow ?? 1;

    /// <InheritDoc />
    public async IAsyncEnumerable<IRowAsync?> GetRowDataAsync(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default)
    {
        await CheckLocationAsync(startRow, ct).ConfigureAwait(false);
        while (_sheetReader!.CurrentRow < SheetDimensions.Height)
        {
            yield return await _sheetReader.GetNextRowAsync(cellGetMode, ct).ConfigureAwait(false);
        }
    }

    /// <inheritdoc/>
    public IEnumerable<IRow?> GetRowData(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default)
    {
        CheckLocation(startRow, ct);
        while (_sheetReader!.CurrentRow < SheetDimensions.Height
               && !ct.IsCancellationRequested)
        {
            yield return _sheetReader.GetNextRow(cellGetMode, ct);
        }
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<Cell[]?> GetRowDataAsync(int startRow, int excelStartColumn, int excelEndColumn,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        await foreach (IRowAsync? row in GetRowDataAsync(startRow, RowCellGet.None, ct).ConfigureAwait(false))
        {
            if (row is null
                || ct.IsCancellationRequested)
            {
                yield break;
            }
            try
            {
                int length = excelEndColumn - excelStartColumn + 1;
                Cell[] cells = new Cell[length];
                int idx = 0;
                for (int i = excelStartColumn; i <= excelEndColumn; i++)
                {
                    cells[idx++] = await row.GetCellAsync(i, ct).ConfigureAwait(false);
                }

                yield return cells;
            }
            finally
            {
                row.Dispose();
            }
        }
    }

    /// <InheritDoc />
    public IEnumerable<Cell[]?> GetRowData(int startRow, int excelStartColumn, int excelEndColumn,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        foreach (IRow? row in GetRowData(startRow, RowCellGet.None, ct))
        {
            if (row is null
                || ct.IsCancellationRequested)
            {
                yield break;
            }

            try
            {
                int length = excelEndColumn - excelStartColumn + 1;
                Cell[] cells = new Cell[length];
                int idx = 0;
                for (int i = excelStartColumn; i <= excelEndColumn; i++)
                {
                    cells[idx++] = row.GetCell(i, ct);
                }

                yield return cells;
            }
            finally
            {
                row.Dispose();
            }
        }
    }

    /// <inheritdoc/>
    public IAsyncEnumerable<Cell[]?> GetRowDataAsync(int startRow, ReadOnlySpan<char> startExcelColumn,
        ReadOnlySpan<char> endExcelColumn, CancellationToken ct = default)
    {
        int excelStartColumn = startExcelColumn.IntParse();
        int excelEndColumn = endExcelColumn.IntParse();
        return GetRowDataAsync(startRow, excelStartColumn, excelEndColumn, ct);
    }

    /// <inheritdoc/>
    public IEnumerable<Cell[]?> GetRowData(int startRow, ReadOnlySpan<char> startExcelColumn,
        ReadOnlySpan<char> endExcelColumn, CancellationToken ct = default)
    {
        int excelStartColumn = startExcelColumn.IntParse();
        int excelEndColumn = endExcelColumn.IntParse();
        return GetRowData(startRow, excelStartColumn, excelEndColumn, ct);
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<Cell[]> GetDefinedRangeAsync(DefinedRange range, [EnumeratorCancellation] CancellationToken ct)
    {
        await foreach (Cell[]? rowCells in GetRowDataAsync(range.ExcelRowStart - 1, range.ExcelColumnStart, range.ExcelColumnEnd, ct).ConfigureAwait(false))
        {
            if (rowCells == null
                || _sheetReader!.CurrentRow > range.ExcelRowEnd)
            {
                yield break;
            }

            yield return rowCells;
        }
    }

    /// <InheritDoc />
    public IEnumerable<Cell[]> GetDefinedRange(DefinedRange range, [EnumeratorCancellation] CancellationToken ct)
    {
        foreach (Cell[]? rowCells in GetRowData(range.ExcelRowStart - 1, range.ExcelColumnStart, range.ExcelColumnEnd, ct))
        {
            if (rowCells == null
                || _sheetReader!.CurrentRow > range.ExcelRowEnd
                || ct.IsCancellationRequested)
            {
                yield break;
            }

            yield return rowCells;
        }
    }

    private async Task CheckLocationAsync(int startRow, [EnumeratorCancellation] CancellationToken ct)
    {
        if (_sheetReader == null
            || _sheetReader.CurrentRow > startRow
           )
        {
            if (_sheetReader != null)
            {
                _sheetReader.Dispose();
                if (!_stream.CanSeek)
                {
                    // TODO: Not pretty, need to sort this to allow multi open, when using Zip stream!
                    throw new NotSupportedException(
                        "Please open sheet with `OverrideOptionsAndUseSheetOnlyOnce = false`; Or, do not attempt to go backwards with an existing sheet instance");
                }
                else
                {
                    _stream.Position = 0;
                }
            }

            _sheetReader = await _xmlSharedReaderHelper.CreateSheetReaderAsync(_stream, _instanceContext, _sharedNameTable, ct).ConfigureAwait(false);
        }
        while (_sheetReader.CurrentRow < startRow)
        {
            IRowAsync? row = await _sheetReader.GetNextRowAsync(RowCellGet.None, ct).ConfigureAwait(false);
            row?.Dispose();
        }
    }

    private void CheckLocation(int startRow, [EnumeratorCancellation] CancellationToken ct)
    {
        if (_sheetReader == null
            || _sheetReader.CurrentRow > startRow
           )
        {
            if (_sheetReader != null)
            {
                _sheetReader.Dispose();
                if (!_stream.CanSeek)
                {
                    // TODO: Not pretty, need to sort this to allow multi open, when using Zip stream!
                    throw new NotSupportedException(
                        "Please open sheet with `OverrideOptionsAndUseSheetOnlyOnce = false`; Or, do not attempt to go backwards with an existing sheet instance");
                }
                else
                {
                    _stream.Position = 0;
                }
            }

            _sheetReader = (IOpenXmlSheetReaderAsync)_xmlSharedReaderHelper.CreateSheetReader(_stream, _instanceContext, _sharedNameTable, ct);
        }
        while (_sheetReader.CurrentRow < startRow
               && !ct.IsCancellationRequested)
        {
            IRow? row = _sheetReader.GetNextRow(RowCellGet.None, ct);
            row?.Dispose();
        }
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _sheetReader?.Dispose();
                _stream.CloseInnerStream();
                _stream.Dispose();
                // Note: _xmlReaderHelper is managed by Excel_PRIME, not by Sheet.
                // Sheet is a consumer of xmlReaderHelper, not its owner.
                // Disposing it here causes issues with shared resources like TempFile
                // that are still referenced by other components (e.g., SharedStrings).
            }

            _isDisposed = true;
        }
    }

    public void Dispose() => Dispose(isDisposing: true);
}