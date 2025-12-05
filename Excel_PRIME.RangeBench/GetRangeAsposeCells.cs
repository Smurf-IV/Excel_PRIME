using System;
using System.Collections.Generic;

using Aspose.Cells;


namespace ExcelPRIME.RangeBench;

public class GetRangeAsposeCells : IGetRange
{
    private Workbook? _workbook;

    public void Dispose()
    {
        _workbook?.Dispose();
        _workbook = null;
    }

    public bool LoadFile(string fullPath)
    {
        _workbook = new Workbook(fullPath, new LoadOptions { CheckDataValid = false, KeepUnparsedData = false, ParsingFormulaOnOpen = false });
        return true;
    }

    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, int? localSheetId = null)
    {
        Aspose.Cells.Range namedRange;
        if (localSheetId.HasValue)
        {
            //Get a specific named range in the worksheet
            namedRange = _workbook.Worksheets.GetRangeByName(definedName, localSheetId.Value, false);
        }
        else
        {
            //Get a specific named range in the workbook
            namedRange = _workbook.Worksheets.GetRangeByName(definedName);
        }
        if (namedRange != null)
        {
            //Get range
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;
            // Iterate through each row in the range
            for (int i = 0; i < rowCount; i++)
            {
                List<object?> rowData = new List<object?>(columnCount);

                // Iterate through each column in the range to access individual cell values
                for (int j = 0; j < columnCount; j++)
                {
                    rowData.Add(namedRange.GetCellOrNull(i, j).Value.ToString());
                }
                yield return rowData;
            }
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName) => throw new NotImplementedException();
}
