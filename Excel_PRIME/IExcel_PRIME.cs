using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;
// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores


/// <summary>
/// This Libraries main Contract
/// </summary>
public interface IExcel_PRIME : IExcelImp
{
    /// <summary>
    /// Opens the file, read-only, and will hold the stream open until disposed
    /// </summary>
    /// <param name="fileName">The full path to the Excel file to be opened.</param>
    /// <param name="options">Optional parameters for configuring the file opening process.</param>
    /// <param name="ct">A <see cref="CancellationToken"/> to observe while waiting for the task to complete.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="fileName"/> is <c>null</c>.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be accessed or opened.</exception>
    /// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel file.</exception>
    void Open(string fileName, Options? options = null, CancellationToken ct = default);

    /// <summary>
    /// _Owns_ the fileStream, until disposed. Must be Seekable.
    /// </summary>
    void Open(Stream fileStream, Options? options = null, CancellationToken ct = default);
}

/// <summary>
/// This Libraries main Contract
/// </summary>
public interface IExcel_PRIMEAsync : IExcel_PRIME, IExcelImpAsync
{
    /// <summary>
    /// Opens the file, read-only, and will hold the stream open until disposed
    /// </summary>
    /// <param name="fileName">The full path to the Excel file to be opened.</param>
    /// <param name="options">Optional parameters for configuring the file opening process.</param>
    /// <param name="ct">A <see cref="CancellationToken"/> to observe while waiting for the task to complete.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="fileName"/> is <c>null</c>.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be accessed or opened.</exception>
    /// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel file.</exception>
    Task OpenAsync(string fileName, Options? options = null, CancellationToken ct = default);

    /// <summary>
    /// _Owns_ the fileStream, until disposed. Must be Seekable.
    /// </summary>
    Task OpenAsync(Stream fileStream, Options? options = null, CancellationToken ct = default);
}

