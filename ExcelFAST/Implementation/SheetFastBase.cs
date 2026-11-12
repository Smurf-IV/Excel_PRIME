using System;
using System.Collections.Generic;
using System.Threading;

using ExcelPRIME;
using ExcelPRIME.Implementation;


namespace ExcelFAST.Implementation;

/// <summary>
/// Magic that makes the sequential recycled memory reader work
/// </summary>
public class SheetFastBase : ISheetFast, IRowFast, IDisposable
{
    protected internal readonly IOpenXmlReaderHelpers _xmlReaderHelper;
    protected internal readonly IZipReader _zipReader;
    protected internal readonly InstanceContext _instanceContext = new();

    private readonly RowFastBase _rowFast;

    private Sheet? _activeSheet;

    protected internal SheetFastBase(IOpenXmlReaderHelpers xmlReader, IZipReader zipReader, IRowFast rowFast)
    {
        _xmlReaderHelper = xmlReader;
        _zipReader = zipReader;
        _rowFast = new RowFastBase(rowFast);
    }

    /// <InheritDoc />
    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(true);
        GC.SuppressFinalize(this);
    }
    private bool _isDisposed;

    /// <summary>
    /// Releases the resources used by this class.
    /// </summary>
    /// <param name="isDisposing">
    /// A value indicating whether the method is being called explicitly 
    /// to release both managed and unmanaged resources (<c>true</c>), 
    /// or by the finalizer to release only unmanaged resources (<c>false</c>).
    /// </param>
    protected virtual void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _rowFast.SetActiveSheet(null);
                _activeSheet?.Dispose();
                _activeSheet = null;
                _instanceContext.SharedStrings?.Dispose();
                _instanceContext.SharedStrings = null;
                _zipReader.Dispose();
                _xmlReaderHelper.Dispose();
            }
            _isDisposed = true;
        }
    }

    /// <InheritDoc />
    public void UseThisSheet(string sheetName, IReadOnlyDictionary<string /*sheetName*/, string /*pathOffsetSheet*/> sheetNamesToPathOffset,
        CancellationToken ct)
    {
        Name = sheetName;
        _activeSheet?.Dispose();
        _activeSheet = null;
        _activeSheet = !sheetNamesToPathOffset.TryGetValue(sheetName, out string? pathOffsetSheet)
                ? null // ($"{sheetName} does not exist");
                : new Sheet(_zipReader.GetEntry(pathOffsetSheet, ct)!, _xmlReaderHelper, sheetName,
                    _instanceContext);

        _rowFast.SetActiveSheet(_activeSheet);
    }

    /// <InheritDoc />
    public bool IsValid => _activeSheet != null;

    /// <InheritDoc />
    public string? Name { get; private set; }

    /// <InheritDoc />
    public (int Height, int Width) SheetDimensions => _activeSheet?.SheetDimensions ?? (0, 0);

    /// <InheritDoc />
    public int CurrentRow => _rowFast.CurrentRow;

    /// <InheritDoc />
    public bool GetNextRowData(CancellationToken ct = default) => _rowFast.GetNextRowData(ct);

    /// <InheritDoc />
    public PooledCellView GetCellAt(int excelColumn, CancellationToken ct = default) => _rowFast.GetCellAt(excelColumn, ct);

    /// <InheritDoc />
    public PooledCellView GetCellAt(ReadOnlySpan<char> excelColumn, CancellationToken ct = default) => _rowFast.GetCellAt(excelColumn, ct);

    /// <InheritDoc />
    public ReadOnlySpan<Cell> GetAllCells(CancellationToken ct = default) => _rowFast.GetAllCells(ct);

    /// <InheritDoc />
    public ReadOnlySpan<Cell> GetCells(int startExcelColumn, int endExcelColumn, CancellationToken ct = default) => _rowFast.GetCells(startExcelColumn, endExcelColumn, ct);

    /// <InheritDoc />
    public ReadOnlySpan<Cell> GetCells(ReadOnlySpan<char> startExcelColumn, ReadOnlySpan<char> endExcelColumn,
        CancellationToken ct = default) => _rowFast.GetCells(startExcelColumn, endExcelColumn, ct);
}
