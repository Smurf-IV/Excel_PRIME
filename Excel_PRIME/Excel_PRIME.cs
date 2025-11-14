using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.Implementation;
using ExcelPRIME.Shared;

using TernaryBool = bool?;


namespace ExcelPRIME;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
/// <summary>
/// Main entry point into the `Excel_PRIME` API's
/// </summary>
public sealed class Excel_PRIME : IExcel_PRIME
{
    private bool _isDisposed;
    private readonly IXmlReaderHelpers _xmlReaderHelper;
    private readonly IZipReader _zipReader;
    private Stream? _fs;
    private readonly Dictionary<string, TempFile> _baseFiles = new();
    private readonly Dictionary<int, TempFile> _sheetFiles = new();
    private IReadOnlyDictionary<string, int> _sheetNamesToOffsetSheetId = new Dictionary<string, int>().AsReadOnly();
    private readonly InstanceContext _instanceContext = new InstanceContext();
    private readonly SemaphoreLocker _locker = new();

    /// <InheritDoc />
    public Excel_PRIME(IXmlReaderHelpers? xmlReader = null, IZipReader? zipReader = null)
    {
        _xmlReaderHelper = xmlReader ?? new XmlReaderHelpers();
        _zipReader = zipReader ?? new ZipReader();
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
        FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None, 0x8000/*64*1024*/, true);
        return OpenAsync(fs, fileType, options, ct);
    }

    /// <InheritDoc />
    public async Task OpenAsync(Stream fileStream, FileType fileType = FileType.Xlsx, Options? options = null, CancellationToken ct = default)
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

        //TempFile workbook_rels = new TempFile("workbook.xml.rels");
        //_baseFiles["xl/_rels/workbook.xml.rels"] = workbook_rels;
        //using (FileStream targetStream = workbook.FileInfo.OpenWrite())
        //{
        //    await _zipReader.CopyToAsync("xl/_rels/workbook.xml.rels", targetStream, ct).ConfigureAwait(false);
        //}

        //        await GetSheetRelationsAsync(workbook_rels, ct).ConfigureAwait(false);

        //_sheets = sheets.Where(x => sheetRelations.ContainsKey(x.RelationId))
        //    .Select(x => new { Sheet = x, ZipEntry = _archive.GetEntry($"xl/{sheetRelations[x.RelationId].Target}") ?? throw new XlsxHelperException($"zip entry not found for {x.SheetName}.") })
        //    .Select(x => new Worksheet(x.Sheet.SheetName, new WorksheetReader(x.ZipEntry!.Open(), _sharedStringLookup)))
        //    .ToArray();

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
            TempFile shareStrings = new TempFile("sharedStrings.xml");
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

    private async Task GetSheetNamesAsync(Stream? workBookStream, CancellationToken ct)
    {
        using IXmlWorkBookReader wbr = await _xmlReaderHelper.CreateWorkBookReaderAsync(workBookStream, ct)
            .ConfigureAwait(false);
        _sheetNamesToOffsetSheetId = wbr.GetSheetNamesAsync(ct).ToBlockingEnumerable(ct).ToDictionary();
    }

    /// <InheritDoc />
    public IEnumerable<string> SheetNames() => _sheetNamesToOffsetSheetId.Keys;

    /// <InheritDoc />
    public IAsyncEnumerable<object?[]> GetDefinedRangeAsync(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default) => throw new NotImplementedException();

    /// <InheritDoc />
    public async Task<ISheet?> GetSheetAsync(string sheetName, TernaryBool OverrideOptionsAndUseSheetOnlyOnce = null, CancellationToken ct = default)
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
