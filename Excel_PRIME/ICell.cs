using System.Collections.Generic;

namespace ExcelPRIME;

/// <summary>
/// The type of cell as indicated by the Excel schema (Not interpreted)
/// </summary>
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
public enum CellType
{
    Unknown,
    Numeric,
    Formula,
    SharedString,   // string placed in the shared table
    InlineString,   // Probably a RichText string
    Boolean,    // 0 or 1 converted to `bool`
    Error,      // Excel error TODO interpret this please.
    Date        // ISO 8601 Format
}
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member

/// <summary>
/// 
/// </summary>
public interface ICell
{
    /// <summary>
    /// Gets the value as read from the file, wrapped in a CellValue.
    /// </summary>
    /// <remarks>
    /// Could be the actual value type if specified, otherwise `string?`
    /// </remarks>
    CellValue CellValue { get; }

    /// <summary>
    /// Returns the type as specified in the Excel file attribute
    /// </summary>
    CellType RawExcelType { get; }

    /// <summary>
    /// The Excel column identifier, e.g. `ABY`
    /// </summary>
    IReadOnlyList<char> ColumnLetters { get; }

    /// <summary>
    /// Excel 1 Based
    /// </summary>
    int ExcelColumnOffset { get; }
}