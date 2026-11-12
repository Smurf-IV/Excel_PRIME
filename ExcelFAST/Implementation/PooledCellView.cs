using System;

using ExcelPRIME;


namespace ExcelFAST.Implementation;

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

    public CellValue Value => _cell.CellValue;
    public CellType RawExcelType => _cell.RawExcelType;
    public int ExcelColumnOffset => _cell.ExcelColumnOffset;
}