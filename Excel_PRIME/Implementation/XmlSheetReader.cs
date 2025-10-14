using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class XmlSheetReader : IXmlSheetReader
{
    private readonly ISharedString _sharedStrings;
    private readonly IEnumerator<XElement> _stepper;
    private bool _isDisposed;
    private readonly int _startRow;

    public XmlSheetReader(XDocument document, ISharedString sharedStrings)
    {
        _sharedStrings = sharedStrings;
        XElement? dimElement = document.Descendants()
            .FirstOrDefault(d => d.Name.LocalName == "dimension");
        string? dim = dimElement?.FirstAttribute?.Value;
        if (dim != null)
        {
            var idx = dim.Split(':');
            _startRow = idx[0].GetRowNumber();
            // Might be an empty sheet (i.e. only "A1")
            SheetDimensions = idx.Length == 1
                ? new ValueTuple<int, int>(1, 1)
                : new ValueTuple<int, int>(idx[1].GetRowNumber(), idx[1].GetExcelColumnNumber());
        }
        else
        {
            SheetDimensions = new ValueTuple<int, int>(0, 0);
        }

        _stepper = document.Descendants()
            .Where(d => d.Name.LocalName == "row")
            .GetEnumerator();
        CurrentRow = 0;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _stepper.Dispose();
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

    public Task<IRow?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        return Task.FromResult(GetNextRow(cellGetMode, ct));
    }

    public IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default)
    {
        CurrentRow++;
        if (CurrentRow < _startRow || !_stepper.MoveNext())
        {
            return (IRow?)null;
        }

        IRow nextRow = new Row(_stepper.Current, _sharedStrings, SheetDimensions.Width);
        if (cellGetMode > RowCellGet.None)
        {
            nextRow.GetAllCells(ct);
        }

        return nextRow;
    }

    public Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct)
    {
        throw new NotImplementedException();
    }
}
