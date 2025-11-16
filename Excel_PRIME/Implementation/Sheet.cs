using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal sealed class Sheet : ISheet
{
    private bool _isDisposed;
    private readonly IXmlReaderHelpers _xmlReaderHelper;
    private readonly InstanceContext _instanceContext;
    private readonly XmlNameTable _sharedNameTable;
    private readonly Stream _stream;
    private IXmlSheetReader? _sheetReader;

    /// <summary>
    /// Get the internal file name of this worksheet
    /// </summary>
    internal static string GetFileName(int index) => $"xl/worksheets/sheet{index}.xml";

    internal Sheet(Stream stream, IXmlReaderHelpers xmlReaderHelper, string name, int index, InstanceContext instanceContext)
    {
        _stream = stream;
        _xmlReaderHelper = xmlReaderHelper;
        _instanceContext = instanceContext;
        _sharedNameTable = new SheetRestrictedNameTable();
        Name = name;
        Index = index;
    }

    /// <InheritDoc />
    public string Name { get; }

    /// <InheritDoc />
    public int Index { get; }

    /// <inheritdoc/>
    public (int Height, int Width) SheetDimensions => _sheetReader!.SheetDimensions;

    /// <inheritdoc/>
    public int CurrentRow => _sheetReader?.CurrentRow ?? 1;

    /// <InheritDoc />
    public async IAsyncEnumerable<IRow?> GetRowDataAsync(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default)
    {
        await CheckLocationAsync(startRow, ct).ConfigureAwait(false);
        while (_sheetReader!.CurrentRow < SheetDimensions.Height)
        {
            yield return await _sheetReader.GetNextRowAsync(cellGetMode, ct).ConfigureAwait(false);
        }
    }

    /// <inheritdoc/>
    public IEnumerable<IRow?> GetRowData(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        CheckLocationAsync(startRow, ct).GetAwaiter().GetResult();
        while (_sheetReader!.CurrentRow < SheetDimensions.Height)
        {
            yield return _sheetReader.GetNextRow(cellGetMode, ct);
        }
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<ICell?[]?> GetRowDataAsync(int startRow, int excelStartColumn, int excelEndColumn, 
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        await foreach (IRow? row in GetRowDataAsync(startRow, RowCellGet.None, ct).ConfigureAwait(false))
        {
            if (row is null)
            {
                yield break;
            }

            List<ICell?> cells = new(excelEndColumn - excelStartColumn + 1);
            for (int i = excelStartColumn; i <= excelEndColumn; i++)
            {
                cells.Add(await row.GetCellAsync(i, ct).ConfigureAwait(false));
            }

            yield return cells.ToArray();
        }
    }

    /// <inheritdoc/>
    public IAsyncEnumerable<ICell?[]?> GetRowDataAsync(int startRow, ReadOnlySpan<char> startExcelColumn,
        ReadOnlySpan<char> endExcelColumn, CancellationToken ct = default)
    {
        int excelStartColumn = startExcelColumn.IntParse();
        int excelEndColumn = endExcelColumn.IntParse();
        return GetRowDataAsync(startRow, excelStartColumn, excelEndColumn, ct);
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<ICell?[]> GetDefinedRangeAsync(DefinedRange range, [EnumeratorCancellation] CancellationToken ct)
    {
        await foreach (ICell?[]? rowCells in GetRowDataAsync(range.ExcelRowStart - 1, range.ExcelColumnStart, range.ExcelColumnEnd, ct).ConfigureAwait(false))
        {
            if (rowCells == null
                || _sheetReader!.CurrentRow > range.ExcelRowEnd)
            {
                yield break;
            }

            yield return rowCells;
        }
    }

    private async Task CheckLocationAsync(int startRow, CancellationToken ct)
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

            _sheetReader = await _xmlReaderHelper.CreateSheetReaderAsync(_stream, _instanceContext, _sharedNameTable, ct).ConfigureAwait(false);
        }
        while (_sheetReader.CurrentRow < startRow)
        {
            await _sheetReader.GetNextRowAsync(RowCellGet.None, ct).ConfigureAwait(false);
        }
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _stream.Dispose();
            }

            _isDisposed = true;
        }
    }

    /// <inheritdoc/>
    ~Sheet()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }
}
