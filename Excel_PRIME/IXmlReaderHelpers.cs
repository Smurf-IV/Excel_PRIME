using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;

namespace ExcelPRIME;

/// <summary>
/// Allow other implementations of Xml readers
/// </summary>
public interface IOpenXmlReaderHelpers : IDisposable
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
    IOpenXmlWorkBookReader CreateWorkBookReader(IZipReader zipReader, CancellationToken ct);

    /// <summary>
    /// Even tho this is not the Async, please create the Async class
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="instanceContext"></param>
    /// <param name="sharedNameTable"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    IOpenXmlSheetReader CreateSheetReader(NonClosingStream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct);

    /// <summary>
    /// Extracts all cell styles from the workbook's styles.#### file.
    /// </summary>
    /// <remarks>
    /// Per ECMA-376 specification, default/implicit styles are pre-populated with standard IDs.
    /// </remarks>
    /// <returns>
    /// A dictionary mapping style IDs to CellStyle objects. Always includes built-in default styles.
    /// Returns empty/default styles if styles.#### is not found.
    /// </returns>
    IReadOnlyDictionary<int, CellStyle> GetExtractStyles(IZipReaderAsync zipReader, CancellationToken ct);
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
    Task<IOpenXmlWorkBookReaderAsync> CreateWorkBookReaderAsync(IZipReaderAsync zipReader, CancellationToken ct);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="stream">This is _not_ owned by the `IXmlWorkBookReader`</param>
    /// <param name="instanceContext"></param>
    /// <param name="sharedNameTable"></param>
    /// <param name="ct"></param>
    /// <returns></returns>
    Task<IOpenXmlSheetReaderAsync> CreateSheetReaderAsync(NonClosingStream stream, InstanceContext instanceContext,
        XmlNameTable sharedNameTable, CancellationToken ct);

    /// <summary>
    /// Extracts all cell styles from the workbook's styles.#### file.
    /// </summary>
    /// <remarks>
    /// Per ECMA-376 specification, default/implicit styles are pre-populated with standard IDs.
    /// </remarks>
    /// <returns>
    /// A dictionary mapping style IDs to CellStyle objects. Always includes built-in default styles.
    /// Returns empty/default styles if styles.#### is not found.
    /// </returns>
    Task<IReadOnlyDictionary<int, CellStyle>> GetExtractStylesAsync(IZipReaderAsync zipReader, CancellationToken ct);
}