using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_PRIME;

public interface IXmlReaderHelpers
{
    Task<IXmlReader> CreateReaderAsync(Stream stream, CancellationToken ct = default);
}