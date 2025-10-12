using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

public interface IZipReader : IDisposable
{
    /// <summary>
    /// Initializes an instance of the internal ZipReader on the given stream.
    /// </summary>
    /// <param name="fileStream">The input seekable stream.</param>
    /// <remarks>stream is _not_ owned by the zip Archive</remarks>
    Task OpenArchiveAsync(Stream fileStream, System.Threading.CancellationToken ct);

    /// <summary>
    /// Opens the entry (If exists), And copy's (Async) to the supplied stream
    /// </summary>
    /// <returns>true if exists</returns>
    Task<bool> CopyToAsync(string entryName, Stream targetStream, CancellationToken ct);

}
