using System;


namespace ExcelFAST;

/// <summary>
/// Access contract for sheets
/// </summary>
public interface ISheetFast : IRowFast, IDisposable
{
    /// <summary>
    /// Is there a usable sheet name Loaded
    /// </summary>
    bool IsValid { get; }

    /// <summary>
    /// This Sheets name
    /// </summary>
    string? Name { get; }

    /// <summary>
    /// What are the Max dimension defined [Excel Rows, Excel Cells] (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }
}
