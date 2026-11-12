using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

using ExcelPRIME;


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
    private IReadOnlyDictionary<string /*sheetName*/, string /*pathOffsetSheet*/> _sheetNamesToPathOffset = new Dictionary<string, string>(StringComparer.Ordinal);

    //private IReadOnlyDictionary<string, DefinedRange>? _definedRanges;

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

    ///// <InheritDoc />
    //public IEnumerable<CellValue[]> GetDefinedRange(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default)
    //{
    //    if (_definedRanges == null)
    //    {
    //        // Lazy load on first use
    //        using IOpenXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(_zipReader, ct);
    //        _definedRanges = wbr.GetDefinedRanges(_sheetNamesToPathOffset, ct);
    //    }

    //    DefinedRange? definedRange = null;
    //    // Perhaps Caller is trying to use a localSheetId reference via `useThisSheetName`
    //    if (!string.IsNullOrEmpty(useThisSheetName))
    //    {
    //        _definedRanges.TryGetValue(string.Concat(rangeName, " (", useThisSheetName, ")"), out definedRange);
    //    }
    //    // Maybe it is not an override of the `localSheetId`, so try the expected reference
    //    if (definedRange == null)
    //    {
    //        if (!_definedRanges.TryGetValue(rangeName, out definedRange))
    //        {
    //            throw new KeyNotFoundException(
    //                $"rangeName: [{rangeName}] and useThisSheetName :[{useThisSheetName}] combo not found");
    //        }
    //    }

    //    if (definedRange.ConstValue != null)
    //    {
    //        yield return [CellValue.Create(definedRange.ConstValue, -1)];
    //        yield break;
    //    }

    //    string? definedRangeSheetName = useThisSheetName ?? definedRange.SheetName;
    //    if (string.IsNullOrEmpty(definedRangeSheetName))
    //    {
    //        // range might be the following definition
    //        // <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
    //        // Or user has made a mistake
    //        yield break;
    //    }

    //    using ISheet? targetSheet = GetSheet( definedRangeSheetName, false, ct);
    //    if (targetSheet == null)
    //    {
    //        throw new KeyNotFoundException($"{definedRangeSheetName} does not exist");
    //    }

    //    foreach (Cell[] rowCells in targetSheet.GetDefinedRange(definedRange, ct))
    //    {
    //        yield return rowCells.Select(cell => cell.CellValue).ToArray();
    //    }
    //}


    ///// <InheritDoc />
    //public IEnumerable<CellValue[]> GetUserRange(string range, string sheetName, [EnumeratorCancellation] CancellationToken ct = default)
    //{
    //    using ISheet? targetSheet = GetSheet(sheetName, ct: ct);
    //    if (targetSheet == null)
    //    {
    //        throw new KeyNotFoundException($"{sheetName} does not exist");
    //    }
    //    DefinedRange definedRange = new(range, sheetName);

    //    foreach (Cell[] rowCells in targetSheet.GetDefinedRange(definedRange, ct))
    //    {
    //        yield return rowCells.Select(cell => cell.CellValue).ToArray();
    //    }
    //}

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