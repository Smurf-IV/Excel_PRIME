using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

public interface IXmlReaderHelpers
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
    Task<IXmlWorkBookReader?> CreateWorkBookReaderAsync(Stream stream, CancellationToken ct);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="sharedStrings"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    Task<IXmlSheetReader?> CreateSheetReaderAsync(Stream stream, ISharedString sharedStrings, CancellationToken ct);
}