using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;

using ExcelFAST.XlsbImp;
using ExcelFAST.XlsxImp;

using ExcelPRIME;
using ExcelPRIME.Implementation;


namespace ExcelFAST.Implementation;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
/// <summary>
/// Main entry point into the `Excel_PRIME` XLSX handler API's
/// </summary>
public class ExcelFastBase : SheetFastBase, IExcelFast
{
    private bool _isDisposed;
    private Stream? _fs;
    /// <summary>
    /// Sheet names mapped to their entry path in the archive.
    /// </summary>
    protected internal IReadOnlyDictionary<string /*sheetName*/, string /*pathOffsetSheet*/> _sheetNamesToPathOffset = new Dictionary<string, string>(StringComparer.Ordinal);

    private IReadOnlyDictionary<string, DefinedRange>? _definedRanges;

    /// <summary>
    /// Loads defined ranges from the workbook.
    /// </summary>
    public virtual IReadOnlyDictionary<string, DefinedRange> LoadDefinedRanges(CancellationToken ct)
    {
        using IOpenXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(_zipReader, ct);
        return wbr.GetDefinedRanges(_sheetNamesToPathOffset, ct);
    }

    /// <summary>
    /// Creates a new format-specific row reader for range enumeration.
    /// </summary>
    protected virtual IRowFast CreateRowFast() => new RowFastXlsx();

    /// <InheritDoc />
    protected internal ExcelFastBase(IOpenXmlReaderHelpers xmlReader, IZipReader zipReader, IRowFast rowFast)
    : base(xmlReader, zipReader, rowFast)
    {
    }

    /// <InheritDoc />
    public void Open(string fileName, Options? options = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileName);
        FileStream fs = new(fileName, FileMode.Open, FileAccess.Read, FileShare.Read, 0x8000/*64*1024*/, true);
        Open(fs, options, ct);
    }

    /// <InheritDoc />
    public void Open(Stream fileStream, Options? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileStream);
        _fs = fileStream;
        if (!fileStream.CanSeek)
        {
            throw new EndOfStreamException("'fileStream' _must_ be seekable!");
        }

        options ??= new Options();
        _instanceContext.Options = options;

        _zipReader.OpenArchive(fileStream, ct);
        // Check and get the Shared strings
        GetSharedStrings(ct);

        // Extract styles from the workbook
        if (options.CellConversionType >= CellConversion.ExcelCellType)
        {
            GetStyles(ct);
        }

        // Now perform the Getting of the base data
        GetSheetNames(_zipReader, ct);
    }

    private void GetSharedStrings(CancellationToken ct)
        => _instanceContext.SharedStrings = _xmlReaderHelper.GetSharedStrings(_zipReader, _instanceContext.Options.AccessExcelFileInForwardOnlyMode, ct);

    private void GetStyles(CancellationToken ct)
        => _instanceContext.CellStyles = _xmlReaderHelper.GetExtractStyles(_zipReader, ct);

    private void GetSheetNames(IZipReader zipReader, CancellationToken ct)
    {
        using IOpenXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(zipReader, ct);
        _sheetNamesToPathOffset = wbr.GetSheetNames(ct).ToDictionary(StringComparer.OrdinalIgnoreCase);
    }

    /// <InheritDoc />
    public IEnumerable<string> SheetNames() => _sheetNamesToPathOffset.Keys;

    /// <InheritDoc />
    public bool SelectSheet(string sheetName, int skipRows = 0, CancellationToken ct = default)
    {
        UseThisSheet(sheetName, _sheetNamesToPathOffset, ct);
        if (IsValid)
        {
            if (skipRows > 0)
            {
                for (int i = 0; i < skipRows; i++)
                {
                    if (!GetNextRowData(ct))
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        return false;
    }

    /// <InheritDoc />
    public IEnumerable<PooledCellView[]> GetDefinedRange(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_definedRanges == null)
        {
            // Lazy load on first use
            _definedRanges = LoadDefinedRanges(ct);
        }

        DefinedRange? definedRange = null;
        // Perhaps Caller is trying to use a localSheetId reference via `useThisSheetName`
        if (!string.IsNullOrEmpty(useThisSheetName))
        {
            _definedRanges.TryGetValue(string.Concat(rangeName, " (", useThisSheetName, ")"), out definedRange);
        }
        // Maybe it is not an override of the `localSheetId`, so try the expected reference
        if (definedRange == null)
        {
            if (!_definedRanges.TryGetValue(rangeName, out definedRange))
            {
                throw new KeyNotFoundException(
                    $"rangeName: [{rangeName}] and useThisSheetName :[{useThisSheetName}] combo not found");
            }
        }

        if (definedRange.ConstValue != null)
        {
            yield return [new PooledCellView(new Cell(CellValue.Create(definedRange.ConstValue, -1), 0, CellType.InlineString))];
            yield break;
        }

        string? definedRangeSheetName = useThisSheetName ?? definedRange.SheetName;
        if (string.IsNullOrEmpty(definedRangeSheetName)
            || !_sheetNamesToPathOffset.ContainsKey(definedRangeSheetName))
        {
            // range might be the following definition
            // <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
            // Or user has made a mistake
            yield break;
        }

        if (!_sheetNamesToPathOffset.TryGetValue(definedRangeSheetName, out string? pathOffsetSheet))
        {
            throw new KeyNotFoundException($"{definedRangeSheetName} does not exist");
        }

        foreach (PooledCellView[] rowCells in EnumerateRange(definedRange, pathOffsetSheet, definedRangeSheetName, ct))
        {
            yield return rowCells;
        }
    }

    /// <InheritDoc />
    public IEnumerable<PooledCellView[]> GetUserRange(string range, string sheetName, [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (!_sheetNamesToPathOffset.TryGetValue(sheetName, out string? pathOffsetSheet))
        {
            throw new KeyNotFoundException($"{sheetName} does not exist");
        }

        DefinedRange definedRange = new(range, sheetName);
        foreach (PooledCellView[] rowCells in EnumerateRange(definedRange, pathOffsetSheet, sheetName, ct))
        {
            yield return rowCells;
        }
    }

    private IEnumerable<PooledCellView[]> EnumerateRange(DefinedRange definedRange, string pathOffsetSheet, string sheetName, [EnumeratorCancellation] CancellationToken ct)
    {
        Stream stream = _zipReader.GetEntry(pathOffsetSheet, ct)!;
        using Sheet sheet = new(stream, _xmlReaderHelper, sheetName, _instanceContext);
        _ = sheet.SheetDimensions;

        using IRowFast rowFast = CreateRowFast();
        if (rowFast is RowFastXlsx rowFastXlsx)
        {
            rowFastXlsx.Attach(sheet._sheetReader as XmlSheetReader);
        }
        else if (rowFast is RowFastXlsb rowFastXlsb)
        {
            rowFastXlsb.Attach(sheet._sheetReader as XlsbSheetReader);
        }

        int colStart = definedRange.ExcelColumnStart;
        int colEnd = definedRange.ExcelColumnEnd;
        int length = colEnd - colStart + 1;
        if (length <= 0)
        {
            yield break;
        }

        while (rowFast.GetNextRowData(ct))
        {
            if (ct.IsCancellationRequested || rowFast.CurrentRow > definedRange.ExcelRowEnd)
            {
                yield break;
            }

            if (rowFast.CurrentRow < definedRange.ExcelRowStart)
            {
                continue;
            }

            PooledCellView[] rowViews = new PooledCellView[length];
            for (int i = 0; i < length; i++)
            {
                rowViews[i] = rowFast.GetCellAt(colStart + i, ct);
            }

            yield return rowViews;
        }
    }

    /// <InheritDoc />

    /// <summary>
    /// Releases the resources used by this class.
    /// </summary>
    /// <param name="isDisposing">
    /// A value indicating whether the method is being called explicitly 
    /// to release both managed and unmanaged resources (<c>true</c>), 
    /// or by the finalizer to release only unmanaged resources (<c>false</c>).
    /// </param>
    protected override void Dispose(bool isDisposing)
    {
        base.Dispose(isDisposing);
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _fs?.Dispose();
                _fs = null;
            }
            _isDisposed = true;
        }
    }

    /// <InheritDoc />
    ~ExcelFastBase()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    /// <InheritDoc />
    public new void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
