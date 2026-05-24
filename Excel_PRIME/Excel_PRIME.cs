using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;

using TernaryBool = bool?;


namespace ExcelPRIME;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
/// <summary>
/// Main entry point into the `Excel_PRIME` XLSX handler API's
/// </summary>
public class Excel_PRIME : IExcel_PRIMEAsync
{
    private bool _isDisposed;
    private readonly IOpenXmlReaderHelpersAsync _xmlReaderHelper;
    private readonly IZipReaderAsync _zipReader;
    private Stream? _fs;
    private readonly Dictionary<string /*pathOffsetSheet*/, TempFile> _sheetFiles = [];
    private IReadOnlyDictionary<string /*sheetName*/, string /*pathOffsetSheet*/> _sheetNamesToPathOffset = new Dictionary<string, string>(StringComparer.Ordinal);
    private readonly InstanceContext _instanceContext = new();
    private readonly SemaphoreLocker _locker = new();
    private IReadOnlyDictionary<string, DefinedRange>? _definedRanges;

    /// <InheritDoc />
    public Excel_PRIME(IOpenXmlReaderHelpersAsync? xmlReader = null, IZipReaderAsync? zipReader = null)
    {
        _xmlReaderHelper = xmlReader ?? new XmlReaderHelpersAsync();
        _zipReader = zipReader ?? new ZipReaderAsync();
    }

    /// <summary>
    /// Asynchronously opens an Excel file for processing.
    /// </summary>
    /// <param name="fileName">The full path to the Excel file to be opened.</param>
    /// <param name="options">Optional parameters for configuring the file opening process.</param>
    /// <param name="ct">A <see cref="CancellationToken"/> to observe while waiting for the task to complete.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="fileName"/> is <c>null</c>.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be accessed or opened.</exception>
    /// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel file.</exception>
    public Task OpenAsync(string fileName, Options? options = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileName);
        FileStream fs = new(fileName, FileMode.Open, FileAccess.Read, FileShare.Read, 0x8000/*64*1024*/, true);
        return OpenAsync(fs, options, ct);
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
    public async Task OpenAsync(Stream fileStream, Options? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileStream);
        _fs = fileStream;
        if (!fileStream.CanSeek)
        {
            throw new EndOfStreamException("'fileStream' _must_ be seekable!");
        }

        options ??= new Options();
        _instanceContext.Options = options;

        await _zipReader.OpenArchiveAsync(fileStream, ct).ConfigureAwait(false);
        // Check and get the Shared strings
        await GetSharedStringsAsync(ct).ConfigureAwait(false);

        // Extract styles from the workbook
        if (options.CellConversionType >= CellConversion.ExcelCellStyle)
        {
            await GetStylesAsync(ct).ConfigureAwait(false);
        }

        // Now perform the Getting of the base data
        await GetSheetNamesAsync(_zipReader, ct).ConfigureAwait(false);
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
        if (options.CellConversionType >= CellConversion.ExcelCellStyle)
        {
            GetStyles(ct);
        }

        // Now perform the Getting of the base data
        GetSheetNames(_zipReader, ct);
    }

    private async Task GetSharedStringsAsync(CancellationToken ct)
        => _instanceContext.SharedStrings = await _xmlReaderHelper.GetSharedStringsAsync(_zipReader, _instanceContext.Options.AccessExcelFileInForwardOnlyMode, ct)
                                            .ConfigureAwait(false);

    private void GetSharedStrings(CancellationToken ct)
        => _instanceContext.SharedStrings = _xmlReaderHelper.GetSharedStrings(_zipReader, _instanceContext.Options.AccessExcelFileInForwardOnlyMode, ct);

    private async Task GetStylesAsync(CancellationToken ct)
        => _instanceContext.CellStyles = await _xmlReaderHelper.GetExtractStylesAsync(_zipReader,ct).ConfigureAwait(false);

    private void GetStyles(CancellationToken ct)
        => _instanceContext.CellStyles = _xmlReaderHelper.GetExtractStyles(_zipReader, ct);

    private async Task GetSheetNamesAsync(IZipReaderAsync zipReader, CancellationToken ct)
    {
        using IOpenXmlWorkBookReaderAsync wbr = await _xmlReaderHelper.CreateWorkBookReaderAsync(zipReader, ct)
            .ConfigureAwait(false);
        _sheetNamesToPathOffset = wbr.GetSheetNamesAsync(ct).ToBlockingEnumerable(ct).ToDictionary(StringComparer.OrdinalIgnoreCase);
    }

    private void GetSheetNames(IZipReader zipReader, CancellationToken ct)
    {
        using IOpenXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(zipReader, ct);
        _sheetNamesToPathOffset = wbr.GetSheetNames(ct).ToDictionary(StringComparer.OrdinalIgnoreCase);
    }

    /// <InheritDoc />
    public IEnumerable<string> SheetNames() => _sheetNamesToPathOffset.Keys;

    /// <InheritDoc />
    public virtual async IAsyncEnumerable<CellValue?[]> GetDefinedRangeAsync(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_definedRanges == null)
        {
            // Lazy load on first use
            using IOpenXmlWorkBookReaderAsync wbr = await _xmlReaderHelper.CreateWorkBookReaderAsync(_zipReader, ct)
                .ConfigureAwait(false);
            _definedRanges = await wbr.GetDefinedRangesAsync(_sheetNamesToPathOffset, ct).ConfigureAwait(false);
        }

        if (!_definedRanges.TryGetValue(rangeName, out DefinedRange? definedRange))
        {
            yield break;
        }

        if (definedRange.ConstValue != null)
        {
            yield return [CellValue.Create(definedRange.ConstValue, -1)];
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

        using ISheetAsync? targetSheet = await GetSheetAsync(definedRangeSheetName, false, ct).ConfigureAwait(false);
        if (targetSheet == null)
        {
            throw new KeyNotFoundException($"{definedRangeSheetName} does not exist");
        }

        await foreach (ICell?[] rowCells in targetSheet.GetDefinedRangeAsync(definedRange, ct).ConfigureAwait(false))
        {
            yield return rowCells.Select(cell => cell?.CellValue).ToArray();
        }
    }

    /// <InheritDoc />
    public IEnumerable<CellValue?[]> GetDefinedRange(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_definedRanges == null)
        {
            // Lazy load on first use
            using IOpenXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(_zipReader, ct);
            _definedRanges = wbr.GetDefinedRanges(_sheetNamesToPathOffset, ct);
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
            yield return [CellValue.Create(definedRange.ConstValue, -1)];
            yield break;
        }

        string? definedRangeSheetName = useThisSheetName ?? definedRange.SheetName;
        if (string.IsNullOrEmpty(definedRangeSheetName))
        {
            // range might be the following definition
            // <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
            // Or user has made a mistake
            yield break;
        }

        using ISheet? targetSheet = GetSheet( definedRangeSheetName, false, ct);
        if (targetSheet == null)
        {
            throw new KeyNotFoundException($"{definedRangeSheetName} does not exist");
        }

        foreach (ICell?[] rowCells in targetSheet.GetDefinedRange(definedRange, ct))
        {
            yield return rowCells.Select(cell => cell?.CellValue).ToArray();
        }
    }


    /// <InheritDoc />
    public async IAsyncEnumerable<CellValue?[]> GetUserRangeAsync(string range, string sheetName, [EnumeratorCancellation] CancellationToken ct = default)
    {
        using ISheetAsync? targetSheet = await GetSheetAsync(sheetName, ct: ct).ConfigureAwait(false);
        if (targetSheet == null)
        {
            throw new KeyNotFoundException($"{sheetName} does not exist");
        }
        DefinedRange definedRange = new DefinedRange(range, sheetName);

        await foreach (ICell?[] rowCells in targetSheet.GetDefinedRangeAsync(definedRange, ct).ConfigureAwait(false))
        {
            yield return rowCells.Select(cell => cell?.CellValue).ToArray();
        }
    }

    /// <InheritDoc />
    public IEnumerable<CellValue?[]> GetUserRange(string range, string sheetName, [EnumeratorCancellation] CancellationToken ct = default)
    {
        using ISheet? targetSheet = GetSheet(sheetName, ct: ct);
        if (targetSheet == null)
        {
            throw new KeyNotFoundException($"{sheetName} does not exist");
        }
        DefinedRange definedRange = new DefinedRange(range, sheetName);

        foreach (ICell?[] rowCells in targetSheet.GetDefinedRange(definedRange, ct))
        {
            yield return rowCells.Select(cell => cell?.CellValue).ToArray();
        }
    }

    /// <InheritDoc />
    public async Task<ISheetAsync?> GetSheetAsync(string sheetName, TernaryBool overrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default)
    {
        if (!_sheetNamesToPathOffset.TryGetValue(sheetName, out string? pathOffsetSheet))
        {
            return null; // ($"{sheetName} does not exist");
        }

        Stream stream;
        if (!overrideOptionsAndUseSheetOnlyOnce.GetValueOrDefault(true)
            && !_instanceContext.Options.AccessExcelFileInForwardOnlyMode
           )
        {
            TempFile sheetFile = await _locker.LockAsync(async () =>
            {
                if (!_sheetFiles.TryGetValue(pathOffsetSheet, out TempFile? sheetFile))
                {
                    sheetFile = new TempFile(Path.GetFileName(pathOffsetSheet));
                    _sheetFiles[pathOffsetSheet] = sheetFile;
                    using FileStream targetStream = sheetFile.OpenForAsyncWrite();
                    await _zipReader.CopyToAsync(pathOffsetSheet, targetStream, ct).ConfigureAwait(false);
                }

                return sheetFile;
            }).ConfigureAwait(false);
            stream = sheetFile.OpenForAsyncRead(true);
        }
        else
        {
            stream = _zipReader.GetEntry(pathOffsetSheet)!;
        }
        return new Sheet(stream, _xmlReaderHelper, sheetName, _instanceContext);
    }

    /// <InheritDoc />
    public ISheet? GetSheet(string sheetName, TernaryBool overrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default)
    {
        // Find Id
        if (!_sheetNamesToPathOffset.TryGetValue(sheetName, out string? pathOffsetSheet))
        {
            return null; // ($"{sheetName} does not exist");
        }

        Stream stream;
        if (!overrideOptionsAndUseSheetOnlyOnce.GetValueOrDefault(true)
            && !_instanceContext.Options.AccessExcelFileInForwardOnlyMode
           )
        {
            TempFile? sheetFile = null;
            _locker.Lock(() =>
            {
                if (!_sheetFiles.TryGetValue(pathOffsetSheet, out sheetFile))
                {
                    sheetFile = new TempFile(Path.GetFileName(pathOffsetSheet));
                    _sheetFiles[pathOffsetSheet] = sheetFile;
                    using FileStream targetStream = sheetFile.OpenForAsyncWrite();
                    _zipReader.CopyTo(pathOffsetSheet, targetStream, ct);
                }
            });
            stream = sheetFile!.OpenForAsyncRead(true);
        }
        else
        {
            stream = _zipReader.GetEntry(pathOffsetSheet)!;
        }
        return new Sheet(stream, _xmlReaderHelper, sheetName, _instanceContext);
    }

    /// <summary>
    /// Releases the resources used by the <see cref="Excel_PRIME"/> class.
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
                _instanceContext.SharedStrings?.Dispose();
                _instanceContext.SharedStrings = null;
                foreach ((string _, TempFile tf) in _sheetFiles)
                {
                    tf.Dispose();
                }
                _zipReader.Dispose();
                _fs?.Dispose();
                _fs = null;
                _locker.Dispose();
            }
            _isDisposed = true;
        }
    }

    /// <InheritDoc />
    ~Excel_PRIME()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    /// <InheritDoc />
    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}