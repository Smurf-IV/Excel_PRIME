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
        return sharedStringsStream == null
            ? new XmlLazyLoadSharedStrings()
            : new XmlLazyLoadSharedStrings(sharedStringsStream, ct);
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
        return sharedStringsStream == null
            ? new XmlLazyLoadSharedStrings()
            : new XmlLazyLoadSharedStrings(sharedStringsStream, ct);
    }


    /// <InheritDoc />
    public async Task<IOpenXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct)
    {
        Stream? stream = await zipReader.GetEntryAsync("xl/workbook.xml", ct).ConfigureAwait(false);
        return new XmlWorkBookReader(stream!, ct);
    }

    /// <InheritDoc />
    public IOpenXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct)
    {
        Stream? stream = zipReader.GetEntry("xl/workbook.xml");
        return new XmlWorkBookReader(stream!, ct);
    }


    /// <InheritDoc />
    public Task<IOpenXmlSheetReaderAsync> CreateSheetReaderAsync(NonClosingStream stream,
        InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct)
    {
        IOpenXmlSheetReaderAsync reader = new XmlSheetReader(stream, instanceContext, sharedNameTable, ct);
        return Task.FromResult(reader);
    }

    public string GetSheetFileName(int offsetSheetId) => $"xl/worksheets/sheet{offsetSheetId}.xml";

    /// <InheritDoc />
    public IOpenXmlSheetReader CreateSheetReader(NonClosingStream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct)
        => new XmlSheetReader(stream, instanceContext, sharedNameTable, ct);

    public void Dispose() => _shareStrings?.Dispose();
}
