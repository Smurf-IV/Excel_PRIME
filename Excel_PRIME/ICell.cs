using System;

namespace ExcelPRIME;

public enum CellType
{
    Unknown,
    Numeric,
    String,
    SharedString,
    InlineString,
    Boolean,
    Error,
    Date
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
    /// Returns the type as specified in the Excel file attribute if specified, otherwise `string`
    /// </summary>
    CellType RawExcelType { get; }

    /// <summary>
    /// The Excel column identifier, e.g. `ABY`
    /// </summary>
    ReadOnlyMemory<char> ColumnLetters { get; }

    /// <summary>
    /// Excel 1 Based
    /// </summary>
    int ExcelColumnOffset { get; }
}