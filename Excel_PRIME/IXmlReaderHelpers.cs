using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelPRIME;

/// <summary>
/// Allow other implementations of Xml readers
/// </summary>
public interface IXmlReaderHelpers
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="stream">This _is_ owned by the `ISharedString`</param>
    /// <param name="ct"></param>
    /// <returns></returns>
    ISharedString GetSharedStrings(Stream stream, CancellationToken ct);

    /// <summary>
    /// Create the interface implementation to get details out of the WorkBook
    /// Even tho this is not the Async, please create the Async class
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="ct"></param>
    IXmlWorkBookReaderAsync CreateWorkBookReader(Stream? stream, CancellationToken ct);

    /// <summary>
    /// Even tho this is not the Async, please create the Async class
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="instanceContext"></param>
    /// <param name="sharedNameTable"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IXmlSheetReaderAsync CreateSheetReader(Stream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct);
}

/// <summary>
/// Allow other implementations of Xml readers
/// </summary>
public interface IXmlReaderHelpersAsync : IXmlReaderHelpers
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="stream">This _is_ owned by the `ISharedString`</param>
    /// <param name="ct"></param>
    /// <returns></returns>
    Task<ISharedString> GetSharedStringsAsync(Stream stream, CancellationToken ct);

    /// <summary>
    /// Create the interface implementation to get details out of the WorkBook
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="ct"></param>
    Task<IXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(Stream? stream, CancellationToken ct);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="instanceContext"></param>
    /// <param name="sharedNameTable"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    Task<IXmlSheetReaderAsync> CreateSheetReaderAsync(Stream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct);
}