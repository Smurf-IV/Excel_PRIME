using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelPRIME;

/// <summary>
/// Allow other implementations of Xml readers
/// </summary>
public interface IOpenXmlReaderHelpers
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="zipReader"></param>
    /// <param name="optionsAccessExcelFileInForwardOnlyMode"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    ISharedString GetSharedStrings(IZipReader zipReader, bool optionsAccessExcelFileInForwardOnlyMode,
        CancellationToken ct);

    /// <summary>
    /// Create the interface implementation to get details out of the WorkBook
    /// Even tho this is not the Async, please create the Async class
    /// </summary>
    IXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct);

    /// <summary>
    /// Even tho this is not the Async, please create the Async class
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="instanceContext"></param>
    /// <param name="sharedNameTable"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IXmlSheetReader CreateSheetReader(Stream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct);
}

/// <summary>
/// Allow other implementations of Xml readers
/// </summary>
public interface IOpenXmlReaderHelpersAsync : IOpenXmlReaderHelpers
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="zipReader"></param>
    /// <param name="optionsAccessExcelFileInForwardOnlyMode"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    Task<ISharedString> GetSharedStringsAsync(IZipReaderAsync zipReader, bool optionsAccessExcelFileInForwardOnlyMode,
        CancellationToken ct);

    /// <summary>
    /// Create the interface implementation to get details out of the WorkBook
    /// </summary>
    Task<IXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct);

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