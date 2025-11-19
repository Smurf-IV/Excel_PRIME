using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

/// <summary>
/// How to extract the data from the Excel file
/// </summary>
public interface IZipReader : IDisposable
{
    /// <summary>
    /// Initializes an instance of the internal ZipReader on the given stream.
    /// </summary>
    /// <remarks>Seekable stream is _not_ owned by the zip Archive</remarks>
    void OpenArchive(Stream fileStream, CancellationToken ct);

    /// <summary>
    /// Opens the entry (If exists), And copies to the supplied stream
    /// </summary>
    /// <returns>true if exists</returns>
    bool CopyTo(string entryName, Stream targetStream, CancellationToken ct);

    /// <summary>
    /// Helper function to get the actual internal Zip stream of an entry
    /// </summary>
    Stream? GetEntry(string entryName);
}

/// <summary>
/// How to extract the data from the Excel file
/// </summary>
public interface IZipReaderAsync : IZipReader
{
    /// <summary>
    /// Initializes an instance of the internal ZipReader on the given stream.
    /// </summary>
    /// <remarks>Seekable stream is _not_ owned by the zip Archive</remarks>
    Task OpenArchiveAsync(Stream fileStream, CancellationToken ct);

    /// <summary>
    /// Opens the entry (If exists), And copies (Async) to the supplied stream
    /// </summary>
    /// <returns>true if exists</returns>
    Task<bool> CopyToAsync(string entryName, Stream targetStream, CancellationToken ct);
}