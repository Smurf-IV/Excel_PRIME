using System.Collections.Generic;

namespace ExcelPRIME.RangeBench;

public class GetRangeExcelPrime : IGetRange
{
    private readonly Excel_PRIME _workbook = new();

    public void Dispose()
    {
        _workbook.Dispose();
    }

    public bool LoadFile(string fullPath)
    {
        _workbook.Open(fullPath, options: new Options { AccessExcelFileInForwardOnlyMode = false });
        return true;
    }


    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, int? localSheetId = null)
    {
        IEnumerable<object?[]> rangeRows = localSheetId.HasValue
                ? _workbook.GetDefinedRange(definedName, localSheetId.Value)
                : _workbook.GetDefinedRange(definedName);
        foreach (var rangeRow in rangeRows)
        {
            yield return rangeRow;
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName)
    {
        IEnumerable<object?[]> rangeRows = _workbook.GetDefinedRange(userRange, sheetName);
        foreach (var rangeRow in rangeRows)
        {
            yield return rangeRow;
        }
    }

}
