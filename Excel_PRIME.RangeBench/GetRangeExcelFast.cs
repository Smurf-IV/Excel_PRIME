using System.Collections.Generic;
using System.Linq;

using ExcelFAST;
using ExcelFAST.Implementation;

namespace ExcelPRIME.RangeBench;

public class GRExcelFastXlsx : IGetRange
{
    private readonly ExcelFastXlsx _workbook = new();

    public void Dispose() => _workbook.Dispose();

    public bool LoadFile(string fullPath)
    {
        _workbook.Open(fullPath, options: new Options { AccessExcelFileInForwardOnlyMode = false });
        return true;
    }


    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, string? sheetName = null)
    {
        IEnumerable<PooledCellView[]> rangeRows = _workbook.GetDefinedRange(definedName, sheetName);
        foreach (PooledCellView[] rangeRow in rangeRows)
        {
            yield return rangeRow.Select(cell => (object?)cell.ToString());
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName)
    {
        IEnumerable<PooledCellView[]> rangeRows = _workbook.GetDefinedRange(userRange, sheetName);
        foreach (PooledCellView[] rangeRow in rangeRows)
        {
            yield return rangeRow.Select(cell => (object?)cell.ToString());
        }
    }

}
