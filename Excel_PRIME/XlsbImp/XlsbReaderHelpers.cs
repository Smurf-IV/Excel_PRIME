using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;


namespace ExcelPRIME.XlsbImp;

internal sealed class XlsbReaderHelpersAsync : IOpenXmlReaderHelpersAsync
{
    private TempFile? _shareStrings;
    public void Dispose()
    {
        // Dispose of the TempFile if it has been initialized
        _shareStrings?.Dispose();
        _shareStrings = null;
    }
    
    /// <InheritDoc />
    public async Task<ISharedString> GetSharedStringsAsync(IZipReaderAsync zipReader,
        bool optionsAccessExcelFileInForwardOnlyMode, CancellationToken ct)
    {
        Stream? sharedStringsStream = null;
        if (optionsAccessExcelFileInForwardOnlyMode)
        {
            sharedStringsStream = await zipReader.GetEntryAsync("xl/sharedStrings.bin", ct).ConfigureAwait(false);
        }
        else
        {
            _shareStrings = new TempFile("sharedStrings.bin");
            bool exists;
            using (FileStream targetStream = _shareStrings.OpenForAsyncWrite())
            {
                exists = await zipReader.CopyToAsync("xl/sharedStrings.bin", targetStream, ct).ConfigureAwait(false);
            }

            if (exists)
            {
#pragma warning disable CA2000
                sharedStringsStream = _shareStrings.OpenForAsyncRead();
#pragma warning restore CA2000
            }
        }

        // Check that the shared string actually exists
        if (sharedStringsStream == null)
        {
            return new XlsbLazyLoadSharedStrings();
        }

        XlsbLazyLoadSharedStrings sharedStrings = new(sharedStringsStream);
        await sharedStrings.InitializeAsync(ct).ConfigureAwait(false);
        return sharedStrings;
    }

    /// <InheritDoc />
    public ISharedString GetSharedStrings(IZipReader zipReader, bool optionsAccessExcelFileInForwardOnlyMode, CancellationToken ct)
    {
        Stream? sharedStringsStream = null;
        if (optionsAccessExcelFileInForwardOnlyMode)
        {
            sharedStringsStream = zipReader.GetEntry("xl/sharedStrings.bin");
        }
        else
        {
            _shareStrings = new TempFile("sharedStrings.bin");
            bool exists;
            using (FileStream targetStream = _shareStrings.OpenForAsyncWrite())
            {
                exists = zipReader.CopyTo("xl/sharedStrings.bin", targetStream, ct);
            }

            if (exists)
            {
                sharedStringsStream = _shareStrings.OpenForAsyncRead();
            }
        }

        // Check that the shared string actually exists
        if (sharedStringsStream == null)
        {
            return new XlsbLazyLoadSharedStrings();
        }

        XlsbLazyLoadSharedStrings sharedStrings = new(sharedStringsStream);
        sharedStrings.Initialize(ct);
        return sharedStrings;
    }


    /// <InheritDoc />
    public async Task<IOpenXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct)
    {
        XlsbWorkBookReaderAsync xlsbWorkBookReaderAsync = new XlsbWorkBookReaderAsync(zipReader);
        await xlsbWorkBookReaderAsync.InitializeAsync(ct).ConfigureAwait(false);
        return xlsbWorkBookReaderAsync;
    }

    /// <InheritDoc />
    public IOpenXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct) => new XlsbWorkBookReader(zipReader, ct);


    /// <InheritDoc />
    public async Task<IOpenXmlSheetReaderAsync> CreateSheetReaderAsync(NonClosingStream stream,
        InstanceContext instanceContext,
        XmlNameTable _, CancellationToken ct)
    {
        // For modern hardware in 2025, 65536(64KB) is the standard "sweet spot" for many workloads
        XlsbSheetReader reader = new(new BufferedStream(stream, 64 * 1024), instanceContext);
        await reader.InitializeAsync(ct).ConfigureAwait(false);
        return reader;
    }

    public async Task<IReadOnlyDictionary<short, CellStyle>> GetExtractStylesAsync(IZipReaderAsync zipReader,
        CancellationToken ct)
    {
        using XlsbStylesExtractor extractor = new(zipReader);
        return await extractor.ExtractStylesAsync(ct).ConfigureAwait(false);
    }

    /// <InheritDoc />
    public IOpenXmlSheetReader CreateSheetReader(NonClosingStream stream, InstanceContext instanceContext,
        XmlNameTable _, CancellationToken ct)
    {
        // For modern hardware in 2025, 65536(64KB) is the standard "sweet spot" for many workloads
        XlsbSheetReader reader = new(new BufferedStream(stream, 64 * 1024), instanceContext);
        reader.Initialize(ct);
        return reader;
    }

    public IReadOnlyDictionary<short, CellStyle> GetExtractStyles(IZipReaderAsync zipReader, CancellationToken ct)
    {
        using XlsbStylesExtractor extractor = new(zipReader);
        return extractor.ExtractStyles(ct);
    }
}
