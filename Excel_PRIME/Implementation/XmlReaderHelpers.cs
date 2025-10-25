using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;


namespace ExcelPRIME.Implementation;

internal sealed class XmlReaderHelpers : IXmlReaderHelpers
{
    /// <InheritDoc />
    public Task<ISharedString> GetSharedStringsAsync(Stream stream, CancellationToken ct)
    {
        ISharedString ss = new LazyLoadSharedStrings(stream, ct);
        return Task.FromResult(ss);
    }


    /// <InheritDoc />
    public async Task<IXmlWorkBookReader?> CreateWorkBookReaderAsync(Stream stream, CancellationToken ct)
    {
        XDocument? document = await XDocument.LoadAsync(stream, LoadOptions.None, ct).ConfigureAwait(false);
        return new XmlWorkBookReader(document);
    }

    /// <InheritDoc />
    public Task<IXmlSheetReader> CreateSheetReaderAsync(Stream stream, ISharedString sharedStrings, XmlNameTable sharedNameTable, CancellationToken ct)
    {
        IXmlSheetReader xmlSheetReader = new XmlSheetReader(stream, sharedStrings, sharedNameTable, ct);
        return Task.FromResult(xmlSheetReader);
    }
}
