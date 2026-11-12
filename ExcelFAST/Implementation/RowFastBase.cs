using System;
using System.Threading;

using ExcelFAST.XlsbImp;
using ExcelFAST.XlsxImp;

using ExcelPRIME;
using ExcelPRIME.Implementation;

namespace ExcelFAST.Implementation;

internal class RowFastBase : IRowFast
{
    private readonly IRowFast _rowFast;
    private Sheet? _activeSheet;
    private bool _isDisposed;

    public RowFastBase(IRowFast rowFast)
    {
        _rowFast = rowFast;
    }

    public void Dispose()
    {
        if (!_isDisposed)
        {
            _isDisposed = true;
            _rowFast.Dispose();
            _activeSheet = null;
        }
    }

    public void SetActiveSheet(Sheet? activeSheet)
    {
        _activeSheet = activeSheet;
        if (activeSheet != null)
        {
            _ = activeSheet.SheetDimensions;
        }

        if (_rowFast is RowFastXlsx rowFastXlsx)
        {
            rowFastXlsx.Attach(activeSheet?._sheetReader as XmlSheetReader);
        }
        else if (_rowFast is RowFastXlsb rowFastXlsb)
        {
            rowFastXlsb.Attach(activeSheet?._sheetReader as XlsbSheetReader);
        }
    }

    /// <InheritDoc />
    public int CurrentRow => _rowFast.CurrentRow;

    public bool GetNextRowData(CancellationToken ct = default) => _activeSheet != null && !_isDisposed && _rowFast.GetNextRowData(ct);

    public PooledCellView GetCellAt(int excelColumn, CancellationToken ct = default) => _rowFast.GetCellAt(excelColumn, ct);

    public PooledCellView GetCellAt(ReadOnlySpan<char> excelColumn, CancellationToken ct = default) => _rowFast.GetCellAt(excelColumn, ct);

    public ReadOnlySpan<Cell> GetAllCells(CancellationToken ct = default) => _rowFast.GetAllCells(ct);

    public ReadOnlySpan<Cell> GetCells(int startExcelColumn, int endExcelColumn, CancellationToken ct = default) => _rowFast.GetCells(startExcelColumn, endExcelColumn, ct);

    public ReadOnlySpan<Cell> GetCells(ReadOnlySpan<char> startExcelColumn, ReadOnlySpan<char> endExcelColumn,
        CancellationToken ct = default) => _rowFast.GetCells(startExcelColumn, endExcelColumn, ct);
}
