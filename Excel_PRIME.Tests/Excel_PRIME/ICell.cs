using System;

namespace ExcelPRIME;

public interface ICell : IDisposable
{
    /// <summary>
    /// Gets the value as read from the file
    /// </summary>
    /// <remarks>
    /// Could be the actual type if read from binary file, otherwise `string?`
    /// </remarks>
    object? RawValue { get; }

    /// <summary>
    /// Returns the type as specified in the Excel file attribute
    /// </summary>
    Type RawExcelType { get; }

    /// <summary>
    /// The Excel column identifier, e.g. `ABY`
    /// </summary>
    string ColumnLetters { get; }

    /// <summary>
    /// Excel 1 Based
    /// </summary>
    int ExcelColumnOffset { get; }
}