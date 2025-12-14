using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;


namespace ExcelPRIME.XlsbImp;

internal sealed class XmlReaderHelpersXlsbAsync : IOpenXmlReaderHelpersAsync
{
    private TempFile? _shareStrings;

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
            return new LazyLoadSharedStrings();
        }
        return new LazyLoadSharedStrings(sharedStringsStream, ct);
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
            return new LazyLoadSharedStrings();
        }
        return new LazyLoadSharedStrings(sharedStringsStream, ct);
    }


    /// <InheritDoc />
    public async Task<IXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct)
    {
        Stream? stream = await zipReader.GetEntryAsync("xl/workbook.bin", ct).ConfigureAwait(false);
        return new XmlWorkBookReader(stream!, ct);
    }

    /// <InheritDoc />
    public IXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct)
    {
        Stream? stream = zipReader.GetEntry("xl/workbook.bin");
        return new XmlWorkBookReader(stream!, ct);
    }


    /// <InheritDoc />
    public async IXmlSheetReaderAsync CreateSheetReader(Stream zipReader, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct)
    {
        Stream? stream = await zipReader.GetEntryAsync("xl/workbook.bin", ct).ConfigureAwait(false);
        return new XmlSheetReader(stream!, instanceContext, sharedNameTable, ct);
    }

    /// <InheritDoc />
    public IXmlSheetReaderAsync CreateSheetReader(IZipReader zipReader, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct)
        => new XmlSheetReader(zipReader.GetEntry("xl/workbook.xml")!, instanceContext, sharedNameTable, ct);

}
