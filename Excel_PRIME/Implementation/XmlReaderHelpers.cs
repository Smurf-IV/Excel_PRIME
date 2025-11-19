using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;


namespace ExcelPRIME.Implementation;

internal sealed class XmlReaderHelpersAsync : IXmlReaderHelpersAsync
{
    /// <InheritDoc />
    public Task<ISharedString> GetSharedStringsAsync(Stream stream, CancellationToken ct)
        => Task.FromResult(GetSharedStrings(stream, ct));
    
    /// <InheritDoc />
    public ISharedString GetSharedStrings(Stream stream, CancellationToken ct)
        => new LazyLoadSharedStrings(stream, ct);


    /// <InheritDoc />
    public Task<IXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(Stream? stream, CancellationToken ct)
        => Task.FromResult(CreateWorkBookReader( stream, ct));

    /// <InheritDoc />
    public IXmlWorkBookReaderAsync CreateWorkBookReader(Stream? stream, CancellationToken ct)
        => new XmlWorkBookReader(stream, ct);



    /// <InheritDoc />
    public IXmlSheetReaderAsync CreateSheetReader(Stream stream, InstanceContext instanceContext, XmlNameTable sharedNameTable, CancellationToken ct)
        => new XmlSheetReader(stream, instanceContext, sharedNameTable, ct);

    /// <InheritDoc />
    public Task<IXmlSheetReaderAsync> CreateSheetReaderAsync(Stream stream, InstanceContext instanceContext, XmlNameTable sharedNameTable, CancellationToken ct)
        => Task.FromResult(CreateSheetReader(stream, instanceContext, sharedNameTable, ct));
}
