using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

/// <summary>
/// Sheet internal access contract
/// </summary>
public interface IOpenXmlSheetReader : IDisposable
{

    /// <summary>
    /// What are the Max dimension defined [Excel Rows, Excel Cells] (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }

    /// <summary>
    /// 
    /// </summary>
    int CurrentRow { get; }

    /// <summary>
    /// Get the row(s), and populate the cells via `cellGetMode`
    /// </summary>
    IRow? GetNextRow(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default);
}


/// <summary>
/// Sheet internal access contract
/// </summary>
public interface IOpenXmlSheetReaderAsync : IOpenXmlSheetReader
{
    /// <summary>
    /// Get the row(s), and populate the cells via `cellGetMode`
    /// </summary>
    Task<IRowAsync?> GetNextRowAsync(RowCellGet cellGetMode = RowCellGet.None, CancellationToken ct = default);
}