using System;

using ExcelPRIME;


namespace ExcelFAST.Implementation;

/// <summary>
/// Indicates that an object is pooled and should be returned to the pool after use.
/// </summary>
[AttributeUsage(AttributeTargets.ReturnValue)]
public sealed class PooledObjectAttribute : Attribute { }

/// <summary>
/// A view into a pooled Cell. Do not store; copy data if retention is needed.
/// </summary>
public readonly record struct PooledCellView
{
    private readonly Cell _cell;

    internal PooledCellView(Cell cell)
    {
        _cell = cell;
    }

    /// <summary>
    /// Gets the value as read from the file, wrapped in a CellValue.
    /// </summary>
    /// <remarks>
    /// Could be the actual value type if specified, otherwise `string?`
    /// </remarks>
    public CellValue Value => _cell.CellValue;

    /// <summary>
    /// Returns the type as specified in the Excel file attribute
    /// </summary>
    public CellType RawExcelType => _cell.RawExcelType;

    /// <summary>
    /// Gets the column offset within the Excel file.
    /// </summary>
    public int ExcelColumnOffset => _cell.ExcelColumnOffset;

    /// <summary>
    /// Converts the cell to the internal value string representation.
    /// </summary>
    public override string? ToString() => _cell.ToString();
}