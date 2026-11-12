using System;
using System.Collections.Generic;
using System.Threading;


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
    /// Attempts to switch the usable sheet to the named entry.
    /// </summary>
    void UseThisSheet(string sheetName,
        IReadOnlyDictionary<string /*sheetName*/, string /*pathOffsetSheet*/> sheetNamesToPathOffset,
        CancellationToken ct);

    /// <summary>
    /// What are the Max dimension defined [Excel Rows, Excel Cells] (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }

    ///// <summary>
    ///// Using $A$1:$A$1 style, to return data from: a single cell, a single column, a matrix / table 
    ///// </summary>
    //ArraySegment<Cell> GetDefinedRange(DefinedRange range, [EnumeratorCancellation] CancellationToken ct = default);
}
