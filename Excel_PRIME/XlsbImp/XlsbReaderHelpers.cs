using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;


namespace ExcelPRIME.XlsbImp;

internal sealed class XlsbReaderHelpersAsync : IOpenXmlReaderHelpersAsync, IDisposable
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
        return sharedStringsStream == null
            ? new XlsbLazyLoadSharedStrings()
            : new XlsbLazyLoadSharedStrings(sharedStringsStream, ct);
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
        return sharedStringsStream == null
            ? new XlsbLazyLoadSharedStrings()
            : new XlsbLazyLoadSharedStrings(sharedStringsStream, ct);
    }


    /// <InheritDoc />
    public async Task<IOpenXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct)
    {
        Stream? stream = await zipReader.GetEntryAsync("xl/workbook.bin", ct).ConfigureAwait(false);
        return new XlsbWorkBookReader(stream!, ct);
    }

    /// <InheritDoc />
    public IOpenXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct)
    {
        Stream? stream = zipReader.GetEntry("xl/workbook.bin");
        return new XlsbWorkBookReader(stream!, ct);
    }


    /// <InheritDoc />
    public Task<IOpenXmlSheetReaderAsync> CreateSheetReaderAsync(Stream stream, InstanceContext instanceContext,
        XmlNameTable _, CancellationToken ct)
    {
        IOpenXmlSheetReaderAsync reader = new XlsbSheetReader(stream, instanceContext, ct);
        return Task.FromResult(reader);
    }

    public string GetSheetFileName(int offsetSheetId) => $"xl/worksheets/sheet{offsetSheetId}.bin";

    /// <InheritDoc />
    public IOpenXmlSheetReader CreateSheetReader(Stream stream, InstanceContext instanceContext,
        XmlNameTable _, CancellationToken ct)
        => new XlsbSheetReader(stream, instanceContext, ct);

}
