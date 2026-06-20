using System.Collections.Generic;
using System.Linq;

namespace ExcelPRIME.RangeBench;

public class GRExcelPrime : IGetRange
{
    private readonly Excel_PRIME _workbook = new();

    public void Dispose() => _workbook.Dispose();

    public bool LoadFile(string fullPath)
    {
        _workbook.Open(fullPath, options: new Options { AccessExcelFileInForwardOnlyMode = false });
        return true;
    }


    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, string? sheetName = null)
    {
        IEnumerable<CellValue[]> rangeRows = _workbook.GetDefinedRange(definedName, sheetName);
        foreach (CellValue[] rangeRow in rangeRows)
        {
            yield return rangeRow.Select(cell => (object?)cell.ToString());
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName)
    {
        IEnumerable<CellValue[]> rangeRows = _workbook.GetDefinedRange(userRange, sheetName);
        foreach (CellValue[] rangeRow in rangeRows)
        {
            yield return rangeRow.Select(cell => (object?)cell.ToString());
        }
    }

}
