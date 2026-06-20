using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;


namespace ExcelPRIME.Implementation;

internal sealed class XmlReaderHelpersAsync : IOpenXmlReaderHelpersAsync
{
    private TempFile? _shareStrings;

    /// <InheritDoc />
    public async Task<ISharedString> GetSharedStringsAsync(IZipReaderAsync zipReader,
        bool optionsAccessExcelFileInForwardOnlyMode, CancellationToken ct)
    {
        Stream? sharedStringsStream = null;
        if (optionsAccessExcelFileInForwardOnlyMode)
        {
            sharedStringsStream = await zipReader.GetEntryAsync("xl/sharedStrings.xml", ct).ConfigureAwait(false);
        }
        else
        {
            _shareStrings = new TempFile("sharedStrings.xml");
            bool exists;
            using (FileStream targetStream = _shareStrings.OpenForAsyncWrite())
            {
                exists = await zipReader.CopyToAsync("xl/sharedStrings.xml", targetStream, ct).ConfigureAwait(false);
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
            return new XmlLazyLoadSharedStrings();
        }

        XmlLazyLoadSharedStrings sharedStrings = new(sharedStringsStream);
        await sharedStrings.InitializeAsync(ct).ConfigureAwait(false);
        return sharedStrings;
    }

    /// <InheritDoc />
    public ISharedString GetSharedStrings(IZipReader zipReader, bool optionsAccessExcelFileInForwardOnlyMode, CancellationToken ct)
    {
        Stream? sharedStringsStream = null;
        if (optionsAccessExcelFileInForwardOnlyMode)
        {
            sharedStringsStream = zipReader.GetEntry("xl/sharedStrings.xml");
        }
        else
        {
            _shareStrings = new TempFile("sharedStrings.xml");
            bool exists;
            using (FileStream targetStream = _shareStrings.OpenForAsyncWrite())
            {
                exists = zipReader.CopyTo("xl/sharedStrings.xml", targetStream, ct);
            }

            if (exists)
            {
                sharedStringsStream = _shareStrings.OpenForAsyncRead();
            }
        }

        // Check that the shared string actually exists
        if (sharedStringsStream == null)
        {
            return new XmlLazyLoadSharedStrings();
        }

        XmlLazyLoadSharedStrings sharedStrings = new(sharedStringsStream);
        sharedStrings.Initialize(ct);
        return sharedStrings;
    }


    /// <InheritDoc />
    public async Task<IOpenXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct)
    {
        XmlWorkBookReaderAsync xmlWorkBookReaderAsync = new XmlWorkBookReaderAsync(zipReader);
        await xmlWorkBookReaderAsync.InitializeAsync(ct).ConfigureAwait(false);
        return xmlWorkBookReaderAsync;
    }

    /// <InheritDoc />
    public IOpenXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct) => new XmlWorkBookReader(zipReader, ct);


    public async Task<IOpenXmlSheetReaderAsync> CreateSheetReaderAsync(NonClosingStream stream,
        InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct)
    {
        XmlSheetReader reader = new(stream, instanceContext, sharedNameTable);
        await reader.InitializeAsync(ct).ConfigureAwait(false);
        return reader;
    }

    public async Task<IReadOnlyDictionary<short, CellStyle>> GetExtractStylesAsync(IZipReaderAsync zipReader,
        CancellationToken ct)
    {
        using StylesExtractor extractor = new(zipReader);
        IReadOnlyDictionary<short, CellStyle> extractStylesAsync = await extractor.ExtractStylesAsync(ct).ConfigureAwait(false);
        return extractStylesAsync;
    }

    public IOpenXmlSheetReader CreateSheetReader(NonClosingStream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct)
    {
        XmlSheetReader reader = new(stream, instanceContext, sharedNameTable);
        reader.Initialize(ct);
        return reader;
    }

    public IReadOnlyDictionary<short, CellStyle> GetExtractStyles(IZipReaderAsync zipReader, CancellationToken ct)
    {
        using StylesExtractor extractor = new(zipReader);
        return extractor.ExtractStyles(ct);
    }

    public void Dispose() => _shareStrings?.Dispose();
}
