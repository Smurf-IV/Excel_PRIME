using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelPRIME;

public interface IXmlWorkBookReader : IDisposable
{
    Task<IReadOnlyDictionary<string, int>> GetSheetNamesAsync(CancellationToken ct);

    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct);
}

public interface IXmlSheetReader : IDisposable
{
    /// <summary>
    /// What are the Max dimension defined cells (Many may be blank)
    /// </summary>
    (int Height, int Width) SheetDimensions { get; }

    int CurrentRow { get; }

    Task<IRow?> GetNextRowAsync(CancellationToken ct);

    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct);

}
