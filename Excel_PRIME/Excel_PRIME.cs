using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.Implementation;
using ExcelPRIME.Shared;


namespace ExcelPRIME;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
public sealed class Excel_PRIME : IExcel_PRIME
{
    private bool _isDisposed;
    private readonly IXmlReaderHelpers _xmlReaderHelper;
    private readonly IZipReader _zipReader;
    private Stream? _fs;
    private readonly Dictionary<string, TempFile> _baseFiles = new();
    private readonly Dictionary<int, TempFile> _sheetFiles = new();
    private IReadOnlyDictionary<string, int> _sheetNamesWithrId = new Dictionary<string, int>().AsReadOnly();
    private ISharedString? _sharedStrings;
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
        _fs = fileStream;
        await _zipReader.OpenArchiveAsync(fileStream, ct).ConfigureAwait(false);
        // Check and get the Shared strings
        await GetSharedStringsAsync(ct).ConfigureAwait(false);

        // Now perform the Getting of the base data
        TempFile workbook = new TempFile("workbook.xml");
        _baseFiles["xl/workbook.xml"] = workbook;
        using (FileStream targetStream = workbook.OpenForAsyncWrite())
        {
            await _zipReader.CopyToAsync("xl/workbook.xml", targetStream, ct).ConfigureAwait(false);
        }

        await GetSheetNamesAsync(workbook, ct).ConfigureAwait(false);

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
        // Check that the shared string actually exists
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
            _sharedStrings = await _xmlReaderHelper.GetSharedStringsAsync(fileStream, ct)
                .ConfigureAwait(false);
        }
        else
        {
            _sharedStrings = new LazyLoadSharedStrings();
        }

    }

    private async Task GetSheetNamesAsync(TempFile workbook, CancellationToken ct)
    {
        using FileStream fileStream = workbook.OpenForAsyncRead();
        using IXmlWorkBookReader? wbr = await _xmlReaderHelper.CreateWorkBookReaderAsync(fileStream, ct)
            .ConfigureAwait(false);
        _sheetNamesWithrId = await wbr!.GetSheetNamesAsync(ct).ConfigureAwait(false);
    }

    /// <InheritDoc />
    public IEnumerable<string> SheetNames() => _sheetNamesWithrId.Keys;

    /// <InheritDoc />
    public IAsyncEnumerable<object?[]> GetDefinedRangeAsync(string rangeName, string? useThisSheetName = null, [EnumeratorCancellation] CancellationToken ct = default) => throw new NotImplementedException();

    /// <InheritDoc />
    public async Task<ISheet?> GetSheetAsync(string sheetName, CancellationToken ct = default)
    {
        // Find rId
        if (!_sheetNamesWithrId.TryGetValue(sheetName, out int rId))
        {
            throw new KeyNotFoundException($"{sheetName} doe snot exist");
        }

        TempFile sheetFile = await _locker.LockAsync(async () =>
        {
            if (!_sheetFiles.TryGetValue(rId, out TempFile? sheetFile))
            {
                sheetFile = new TempFile($"sheet{rId}.xml");
                _sheetFiles[rId] = sheetFile;
                using FileStream targetStream = sheetFile.OpenForAsyncWrite();
                string sheetFileName = Sheet.GetFileName(rId);
                await _zipReader.CopyToAsync(sheetFileName, targetStream, ct).ConfigureAwait(false);
            }
            return sheetFile;
        }).ConfigureAwait(false);

        FileStream stream = sheetFile.OpenForAsyncRead(true);
        return new Sheet(stream, _xmlReaderHelper, sheetName, rId, _sharedStrings);
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _sharedStrings?.Dispose();
                _sharedStrings = null;
                foreach ((int _, TempFile tf) in _sheetFiles)
                {
                    tf.Dispose();
                }
                foreach (TempFile tf in _baseFiles.Values)
                {
                    tf.Dispose();
                }
                _baseFiles.Clear();
                _fs?.Dispose();
                _fs = null;
                _locker.Dispose();
            }
            _isDisposed = true;
        }
    }

    // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    ~Excel_PRIME()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
