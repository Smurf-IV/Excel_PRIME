using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_PRIME;

public interface IZipReader : IDisposable
{
    /// <summary>
    /// Initializes an instance of the internal ZipReader on the given stream.
    /// </summary>
    /// <param name="stream">The input seekable stream.</param>
    Task OpenArchiveAsync(Stream fileStream, System.Threading.CancellationToken ct);

    /// <summary>
    /// Opens the entry. And copys to the supplied stream
    /// </summary>
    Task CopyToAsync(string entryName, Stream targteStream, CancellationToken ct);

}
