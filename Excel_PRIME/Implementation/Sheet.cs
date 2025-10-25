using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelPRIME.Implementation;

internal sealed class Sheet : ISheet
{
    private bool _isDisposed;
    private readonly IXmlReaderHelpers _xmlReaderHelper;
    private readonly ISharedString _sharedStrings;
    private readonly XmlNameTable _sharedNameTable;
    private readonly FileStream _stream;
    private IXmlSheetReader? _sheetReader;

    /// <summary>
    /// Get the internal file name of this worksheet
    /// </summary>
    internal static string GetFileName(int index) => $"xl/worksheets/sheet{index}.xml";

    internal Sheet(FileStream stream, IXmlReaderHelpers xmlReaderHelper, string name, int index, ISharedString sharedStrings)
    {
        _stream = stream;
        _stream.Position = 0;
        _xmlReaderHelper = xmlReaderHelper;
        _sharedStrings = sharedStrings;
        _sharedNameTable = new NameTable();
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
            _sheetReader?.Dispose();
            _stream.Position = 0;

            _sheetReader = await _xmlReaderHelper.CreateSheetReaderAsync(_stream, _sharedStrings, _sharedNameTable, ct).ConfigureAwait(false);
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
