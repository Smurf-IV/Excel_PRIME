using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelPRIME.Implementation;

internal sealed class Row : IRow
{
    private readonly XElement _rowElement;
    private readonly IReadOnlyList<string> _sharedStrings;
    private readonly int _maxColumnDimension;
    private bool _isDisposed;
    private int? _rowOffset;
    private Dictionary<int, Cell>? _cells;

    public Row(XElement rowElement, IReadOnlyList<string> sharedStrings, int maxColumnDimension)
    {
        _rowElement = rowElement;
        _sharedStrings = sharedStrings;
        _maxColumnDimension = maxColumnDimension;
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
            }

            _isDisposed = true;
        }
    }

    ~Row()
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

    /// <InheritDoc />
    public int RowOffset
    {
        get
        {
            _rowOffset ??= _rowElement.Attributes("r")
                            .Select(a => int.Parse(a.Value, CultureInfo.InvariantCulture))
                            .FirstOrDefault();
            return _rowOffset.GetValueOrDefault();
        }
    }

    /// <InheritDoc />
    public async IAsyncEnumerable<ICell?> GetAllCellsAsync([EnumeratorCancellation] CancellationToken ct = default)
    {
        await GetCellsAsync(ct).ConfigureAwait(false);
        for (int i = 0; i < _maxColumnDimension; i++)
        {
            _cells.TryGetValue(i, out Cell? found);
            yield return found;
        }
    }

    private Task GetCellsAsync(CancellationToken ct)
    {
        if (_cells != null)
        {
            return Task.CompletedTask;
        }

        _cells = new Dictionary<int, Cell>();
        foreach (XElement cellElement in _rowElement.Elements())
        {
            Cell cell = new Cell(cellElement, _sharedStrings);
            _cells.Add(cell.ExcelColumnOffset - 1, cell);
        }
        return Task.CompletedTask;
    }

    /// <InheritDoc />
    public Task<ICell?> GetCellAsync(int excelColumnIndex, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    /// <InheritDoc />
    public Task<ICell?> GetCellAsync(string columnLetters, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }
}
