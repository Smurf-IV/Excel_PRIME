using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;


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
    public Task<IXmlWorkBookReader> CreateWorkBookReaderAsync(Stream? stream, CancellationToken ct)
    {
        IXmlWorkBookReader xmlWorkBookReader = new XmlWorkBookReader(stream, ct);
        return Task.FromResult(xmlWorkBookReader);
    }

    /// <InheritDoc />
    public Task<IXmlSheetReader> CreateSheetReaderAsync(Stream stream, InstanceContext instanceContext, XmlNameTable sharedNameTable, CancellationToken ct)
    {
        IXmlSheetReader xmlSheetReader = new XmlSheetReader(stream, instanceContext, sharedNameTable, ct);
        return Task.FromResult(xmlSheetReader);
    }
}
