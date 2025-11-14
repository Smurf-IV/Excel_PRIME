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

    public (int Height, int Width) SheetDimensions => _sheetReader.SheetDimensions;

    public int CurrentRow => _sheetReader?.CurrentRow ?? 1;

    /// <InheritDoc />
    public async IAsyncEnumerable<IRow?> GetRowDataAsync(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default)
    {
        await CheckLocationAsync(startRow, ct).ConfigureAwait(false);
        while (_sheetReader.CurrentRow < SheetDimensions.Height)
        {
            yield return await _sheetReader.GetNextRowAsync(cellGetMode, ct).ConfigureAwait(false);
        }
    }

    public IEnumerable<IRow?> GetRowData(int startRow = 0, RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        CheckLocationAsync(startRow, ct).GetAwaiter().GetResult();
        while (_sheetReader.CurrentRow < SheetDimensions.Height)
        {
            yield return _sheetReader.GetNextRow(cellGetMode, ct);
        }
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<IRow?> GetRowDataAsync(int startRow, int startColumn, int numberOfColumns, RowCellGet cellGetMode = RowCellGet.None, [EnumeratorCancellation] CancellationToken ct = default)
    {
        await CheckLocationAsync(startRow, ct).ConfigureAwait(false);
        throw new NotImplementedException();
        while (_sheetReader.CurrentRow < SheetDimensions.Height)
        {
            yield return await _sheetReader.GetNextRowAsync(cellGetMode, ct).ConfigureAwait(false);
        }
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<ICell?[]> GetDefinedRangeAsync(string range, [EnumeratorCancellation] CancellationToken ct)
    {
        int startRow = 0;
        await CheckLocationAsync(startRow, ct).ConfigureAwait(false);
        throw new NotImplementedException();
        yield break;
    }

    /// <InheritDoc />
    public async Task<ICell?> GetRangeCellAsync(string rangeCell, CancellationToken ct)
    {
        int startRow = 0;
        await CheckLocationAsync(startRow, ct).ConfigureAwait(false);
        throw new NotImplementedException();
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

    ~Sheet()
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
