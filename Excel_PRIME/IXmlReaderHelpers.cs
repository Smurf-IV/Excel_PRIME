using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_PRIME;

public interface IXmlReaderHelpers
{
    Task<IXmlWorkBookReader> CreateWorkBookReaderAsync(Stream stream, CancellationToken ct = default);
    Task<IXmlSharedStringsReader> CreateSharedStringsReaderAsync(Stream stream, CancellationToken ct = default);
    Task<IXmlSheetReader> CreateSheetReaderAsync(Stream stream, CancellationToken ct = default);
}