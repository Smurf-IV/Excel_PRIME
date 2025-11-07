namespace ExcelPRIME;

/// <summary>
/// The type of cell as indicated by the Excel schema (Not interpreted)
/// </summary>
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

public interface ICell
{
    /// <summary>
    /// Gets the value as read from the file
    /// </summary>
    /// <remarks>
    /// Could be the actual value type if specified, otherwise `string?`
    /// </remarks>
    object? RawValue { get; }

    /// <summary>
    /// Returns the type as specified in the Excel file attribute
    /// </summary>
    CellType RawExcelType { get; }

    /// <summary>
    /// The Excel column identifier, e.g. `ABY`
    /// </summary>
    char[] ColumnLetters { get; }

    /// <summary>
    /// Excel 1 Based
    /// </summary>
    int ExcelColumnOffset { get; }
}