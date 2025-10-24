using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
public interface IExcel_PRIME : IDisposable
{
    /// <summary>
    /// Opens the file, read-only, and will hold the stream open until disposed
    /// </summary>
    /// <param name="fileName">The full path to the Excel file to be opened.</param>
    /// <param name="fileType">The type of the file to be opened. Defaults to <see cref="FileType.Xlsx"/>.</param>
    /// <param name="options">Optional parameters for configuring the file opening process.</param>
    /// <param name="ct">A <see cref="CancellationToken"/> to observe while waiting for the task to complete.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="fileName"/> is <c>null</c>.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be accessed or opened.</exception>
    /// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel file.</exception>
    Task OpenAsync(string fileName, FileType fileType = FileType.Xlsx, Options? options = null, CancellationToken ct = default);

    /// <summary>
    /// _Owns_ the fileStream, until disposed. Must be Seekable.
    /// </summary>
    Task OpenAsync(Stream fileStream, FileType fileType = FileType.Xlsx, Options? options = null, CancellationToken ct = default);

    /// <summary>
    /// What names exist in this file
    /// </summary>
    IEnumerable<string> SheetNames();

    /// <summary>
    /// Switch functionality to a new sheet
    /// </summary>
    Task<ISheet?> GetSheetAsync(string sheetName, CancellationToken ct = default);

    /// <summary>
    /// From the `definedName`s in the xlsx, use the name to return the range data
    /// </summary>
    /// <param name="rangeName"></param>
    /// <param name="useThisSheetName">If passed in, then check that the range exists in that first, before switching to the global name</param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IAsyncEnumerable<object?[]> GetDefinedRangeAsync(string rangeName, string? useThisSheetName=null, [EnumeratorCancellation] CancellationToken ct = default);
}
