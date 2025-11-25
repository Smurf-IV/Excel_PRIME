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
/// Main entry point into the `Excel_PRIME` API's
/// </summary>
public sealed class Excel_PRIME : IExcel_PRIMEAsync
{
    private bool _isDisposed;
    private readonly IXmlReaderHelpersAsync _xmlReaderHelper;
    private readonly IZipReaderAsync _zipReader;
    private Stream? _fs;
    private readonly Dictionary<string, TempFile> _baseFiles = [];
    private readonly Dictionary<int, TempFile> _sheetFiles = [];
    private IReadOnlyDictionary<string, int> _sheetNamesToOffsetSheetId = new Dictionary<string, int>().AsReadOnly();
    private readonly InstanceContext _instanceContext = new();
    private readonly SemaphoreLocker _locker = new();
    private IReadOnlyDictionary<string, DefinedRange>? _definedRanges;

    /// <InheritDoc />
    public Excel_PRIME(IXmlReaderHelpersAsync? xmlReader = null, IZipReaderAsync? zipReader = null)
    {
        _xmlReaderHelper = xmlReader ?? new XmlReaderHelpersAsync();
        _zipReader = zipReader ?? new ZipReaderAsync();
    }

    /// <summary>
    /// Asynchronously opens an Excel file for processing.
    /// </summary>
    /// <param name="fileName">The full path to the Excel file to be opened.</param>
    /// <param name="fileType">The type of the file to be opened. Defaults to <see cref="FileType.Xlsx"/>.</param>
    /// <param name="options">Optional parameters for configuring the file opening process.</param>
    /// <param name="ct">A <see cref="CancellationToken"/> to observe while waiting for the task to complete.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="fileName"/> is <c>null</c>.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be accessed or opened.</exception>
    /// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel file.</exception>
    public Task OpenAsync(string fileName, FileType fileType = FileType.Xlsx, Options? options = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileName);
        FileStream fs = new(fileName, FileMode.Open, FileAccess.Read, FileShare.None, 0x8000/*64*1024*/, true);
        return OpenAsync(fs, fileType, options, ct);
    }
    /// <InheritDoc />
    public void Open(string fileName, FileType fileType = FileType.Xlsx, Options? options = null,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileName);
        FileStream fs = new(fileName, FileMode.Open, FileAccess.Read, FileShare.None, 0x8000/*64*1024*/, true);
        Open(fs, fileType, options, ct);
    }

    /// <InheritDoc />
    public async Task OpenAsync(Stream fileStream, FileType fileType, Options? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileStream);
        if (!fileStream.CanSeek)
        {
            throw new EndOfStreamException("'fileStream' _must_ be seekable!");
        }

        options ??= new Options();
        _instanceContext.Options = options;

        _fs = fileStream;
        await _zipReader.OpenArchiveAsync(fileStream, ct).ConfigureAwait(false);
        // Check and get the Shared strings
        await GetSharedStringsAsync(ct).ConfigureAwait(false);

        // Now perform the Getting of the base data
        Stream workBookStream = _zipReader.GetEntry("xl/workbook.xml")!;
        await GetSheetNamesAsync(workBookStream, ct).ConfigureAwait(false);
    }

    /// <InheritDoc />
    public void Open(Stream fileStream, FileType fileType, Options? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileStream);
        if (!fileStream.CanSeek)
        {
            throw new EndOfStreamException("'fileStream' _must_ be seekable!");
        }

        options ??= new Options();
        _instanceContext.Options = options;

        _fs = fileStream;
        _zipReader.OpenArchive(fileStream, ct);
        // Check and get the Shared strings
        GetSharedStrings(ct);

        // Now perform the Getting of the base data
        Stream workBookStream = _zipReader.GetEntry("xl/workbook.xml")!;
        GetSheetNames(workBookStream, ct);
    }

    private async Task GetSharedStringsAsync(CancellationToken ct)
    {
        _instanceContext.SharedStrings = new LazyLoadSharedStrings();

        // Check that the shared string actually exists
        if (_instanceContext.Options.AccessExcelFileInForwardOnlyMode)
        {
            Stream? sharedStringsStream = _zipReader.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsStream != null)
            {
                _instanceContext.SharedStrings = await _xmlReaderHelper.GetSharedStringsAsync(sharedStringsStream, ct)
                    .ConfigureAwait(false);
            }
        }
        else
        {
            TempFile shareStrings = new("sharedStrings.xml");
            _baseFiles["xl/sharedStrings.xml"] = shareStrings;
            bool exists;
            using (FileStream targetStream = shareStrings.OpenForAsyncWrite())
            {
                exists = await _zipReader.CopyToAsync("xl/sharedStrings.xml", targetStream, ct).ConfigureAwait(false);
            }

            if (exists)
            {
#pragma warning disable CA2000 // <param name="stream">This _is_ owned by the `ISharedString`</param>
                FileStream fileStream = shareStrings.OpenForAsyncRead();
#pragma warning restore CA2000
                _instanceContext.SharedStrings = await _xmlReaderHelper.GetSharedStringsAsync(fileStream, ct)
                    .ConfigureAwait(false);
            }
        }
    }

    private void GetSharedStrings(CancellationToken ct)
    {
        _instanceContext.SharedStrings = new LazyLoadSharedStrings();

        // Check that the shared string actually exists
        if (_instanceContext.Options.AccessExcelFileInForwardOnlyMode)
        {
            Stream? sharedStringsStream = _zipReader.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsStream != null)
            {
                _instanceContext.SharedStrings = _xmlReaderHelper.GetSharedStrings(sharedStringsStream, ct);
            }
        }
        else
        {
            TempFile shareStrings = new("sharedStrings.xml");
            _baseFiles["xl/sharedStrings.xml"] = shareStrings;
            bool exists;
            using (FileStream targetStream = shareStrings.OpenForAsyncWrite())
            {
                exists = _zipReader.CopyTo("xl/sharedStrings.xml", targetStream, ct);
            }

            if (exists)
            {
#pragma warning disable CA2000 // <param name="stream">This _is_ owned by the `ISharedString`</param>
                FileStream fileStream = shareStrings.OpenForAsyncRead();
#pragma warning restore CA2000
                _instanceContext.SharedStrings = _xmlReaderHelper.GetSharedStrings(fileStream, ct);
            }
        }
    }

    private async Task GetSheetNamesAsync(Stream? workBookStream, CancellationToken ct)
    {
        using IXmlWorkBookReaderAsync wbr = await _xmlReaderHelper.CreateWorkBookReaderAsync(workBookStream, ct)
            .ConfigureAwait(false);
        _sheetNamesToOffsetSheetId = wbr.GetSheetNamesAsync(ct).ToBlockingEnumerable(ct).ToDictionary();
    }

    private void GetSheetNames(Stream? workBookStream, CancellationToken ct)
    {
        using IXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(workBookStream, ct);
        _sheetNamesToOffsetSheetId = wbr.GetSheetNames(ct).ToDictionary();
    }

    /// <InheritDoc />
    public IEnumerable<string> SheetNames() => _sheetNamesToOffsetSheetId.Keys;

    /// <InheritDoc />
    public async IAsyncEnumerable<object?[]> GetDefinedRangeAsync(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_definedRanges == null)
        {
            // Lazy load on first use
            Stream workBookStream = _zipReader.GetEntry("xl/workbook.xml")!;
            using IXmlWorkBookReaderAsync wbr = await _xmlReaderHelper.CreateWorkBookReaderAsync(workBookStream, ct)
                .ConfigureAwait(false);
            _definedRanges = await wbr.GetDefinedRangesAsync(_sheetNamesToOffsetSheetId, ct).ConfigureAwait(false);
        }

        if (!_definedRanges.TryGetValue(rangeName, out DefinedRange? definedRange))
        {
            yield break;
        }

        if (definedRange.ConstValue != null)
        {
            yield return [definedRange.ConstValue];
            yield break;
        }

        string definedRangeSheetName = useThisSheetName ?? definedRange.SheetName ??
            _sheetNamesToOffsetSheetId.FirstOrDefault(kvp => kvp.Value == definedRange.SheetIdReference.IntParse()).Key;
        if (!_sheetNamesToOffsetSheetId.ContainsKey(definedRangeSheetName))
        {
            // range might be the following definition
            // <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
            // Or user has made a mistake
            yield break;
        }

        using ISheetAsync? targetSheet = await GetSheetAsync(definedRangeSheetName, false, ct).ConfigureAwait(false);
        if (targetSheet == null)
        {
            yield break;
        }

        await foreach (ICell?[] rowCells in targetSheet.GetDefinedRangeAsync(definedRange, ct).ConfigureAwait(false))
        {
            yield return rowCells.Select(cell => cell?.RawValue).ToArray();
        }
    }

    /// <InheritDoc />
    public IEnumerable<object?[]> GetDefinedRange(string rangeName, int useLocalSheetId, [EnumeratorCancellation] CancellationToken ct = default)
    {
        string? useThisSheetName = null;
        int valueOffset = useLocalSheetId + 1;
        KeyValuePair<string, int> firstOrDefault = _sheetNamesToOffsetSheetId.FirstOrDefault(kvp => kvp.Value == valueOffset);
        if (!string.IsNullOrEmpty(firstOrDefault.Key))
        {
            useThisSheetName = firstOrDefault.Key;
        }
        return GetDefinedRange(rangeName, useThisSheetName, ct);
    }

    /// <InheritDoc />
    public IEnumerable<object?[]> GetDefinedRange(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default)
    {
        if (_definedRanges == null)
        {
            // Lazy load on first use
            Stream workBookStream = _zipReader.GetEntry("xl/workbook.xml")!;
            using IXmlWorkBookReader wbr = _xmlReaderHelper.CreateWorkBookReader(workBookStream, ct);
            _definedRanges = wbr.GetDefinedRanges(_sheetNamesToOffsetSheetId, ct);
        }

        DefinedRange? definedRange = null;
        // Perhaps Caller is trying to use a localSheetId reference via `useThisSheetName`
        if (!string.IsNullOrEmpty(useThisSheetName))
        {
            _definedRanges.TryGetValue(string.Concat(rangeName, " (", useThisSheetName, ")"), out definedRange);
        }
        // Maybe it is not an override of the `localSheetId`, so try the expected reference
        if ( definedRange == null)
        {
            if (!_definedRanges.TryGetValue(rangeName, out definedRange))
            {
                throw new KeyNotFoundException(
                    $"rangeName: [{rangeName}] and useThisSheetName :[{useThisSheetName}] combo not found");
            }
        }

        if (definedRange.ConstValue != null)
        {
            yield return [definedRange.ConstValue];
            yield break;
        }

        using ISheet? targetSheet = GetSheet(
            useThisSheetName ?? definedRange.SheetName ??
            _sheetNamesToOffsetSheetId.First(kvp => kvp.Value == definedRange.SheetIdReference.IntParse()).Key,
            false, ct);
        if (targetSheet == null)
        {
            yield break;
        }

        foreach (ICell?[] rowCells in targetSheet.GetDefinedRange(definedRange, ct))
        {
            yield return rowCells.Select(cell => cell?.RawValue).ToArray();
        }
    }

    /// <InheritDoc />
    public async Task<ISheetAsync?> GetSheetAsync(string sheetName, TernaryBool OverrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default)
    {
        // Find Id
        if (!_sheetNamesToOffsetSheetId.TryGetValue(sheetName, out int offsetSheetId))
        {
            throw new KeyNotFoundException($"{sheetName} does not exist");
        }

        Stream stream;
        if (!OverrideOptionsAndUseSheetOnlyOnce.GetValueOrDefault(true)
            && !_instanceContext.Options.AccessExcelFileInForwardOnlyMode
           )
        {
            TempFile sheetFile = await _locker.LockAsync(async () =>
            {
                if (!_sheetFiles.TryGetValue(offsetSheetId, out TempFile? sheetFile))
                {
                    sheetFile = new TempFile($"sheet{offsetSheetId}.xml");
                    _sheetFiles[offsetSheetId] = sheetFile;
                    using FileStream targetStream = sheetFile.OpenForAsyncWrite();
                    string sheetFileName = Sheet.GetFileName(offsetSheetId);
                    await _zipReader.CopyToAsync(sheetFileName, targetStream, ct).ConfigureAwait(false);
                }

                return sheetFile;
            }).ConfigureAwait(false);
            stream = sheetFile.OpenForAsyncRead(true);
        }
        else
        {
            string sheetFileName = Sheet.GetFileName(offsetSheetId);
            stream = _zipReader.GetEntry(sheetFileName)!;
        }
        return new Sheet(stream, _xmlReaderHelper, sheetName, offsetSheetId, _instanceContext);
    }

    /// <InheritDoc />
    public ISheet? GetSheet(string sheetName, TernaryBool OverrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default)
    {
        // Find Id
        if (!_sheetNamesToOffsetSheetId.TryGetValue(sheetName, out int offsetSheetId))
        {
            throw new KeyNotFoundException($"{sheetName} does not exist");
        }

        Stream stream;
        if (!OverrideOptionsAndUseSheetOnlyOnce.GetValueOrDefault(true)
            && !_instanceContext.Options.AccessExcelFileInForwardOnlyMode
           )
        {
            TempFile? sheetFile = null;
            _locker.Lock(() =>
            {
                if (!_sheetFiles.TryGetValue(offsetSheetId, out sheetFile))
                {
                    sheetFile = new TempFile($"sheet{offsetSheetId}.xml");
                    _sheetFiles[offsetSheetId] = sheetFile;
                    using FileStream targetStream = sheetFile.OpenForAsyncWrite();
                    string sheetFileName = Sheet.GetFileName(offsetSheetId);
                    _zipReader.CopyTo(sheetFileName, targetStream, ct);
                }
            });
            stream = sheetFile!.OpenForAsyncRead(true);
        }
        else
        {
            string sheetFileName = Sheet.GetFileName(offsetSheetId);
            stream = _zipReader.GetEntry(sheetFileName)!;
        }
        return new Sheet(stream, _xmlReaderHelper, sheetName, offsetSheetId, _instanceContext);
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _instanceContext.SharedStrings?.Dispose();
                _instanceContext.SharedStrings = null;
                foreach ((int _, TempFile tf) in _sheetFiles)
                {
                    tf.Dispose();
                }
                foreach (TempFile tf in _baseFiles.Values)
                {
                    tf.Dispose();
                }
                _zipReader.Dispose();
                _baseFiles.Clear();
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
