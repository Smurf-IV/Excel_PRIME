using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using ClosedXML.Excel;

namespace ExcelPRIME.RangeBench;

public class GRClosedXML : IGetRange
{
    private XLWorkbook? wb;

    public void Dispose()
    {
        wb?.Dispose();
        wb = null;
    }

    public bool LoadFile(string fullPath)
    {
        wb = new XLWorkbook(fullPath);
        return wb != null;
    }

    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, string? sheetName = null)
    {
        IXLRanges rangesLocal;
        // worksheet scope
        if (sheetName != null)
        {
            wb!.Worksheets.TryGetWorksheet(sheetName, out IXLWorksheet? worksheet);
            rangesLocal = worksheet.Ranges(definedName);
        }
        else
        {
            // workbook scope
            rangesLocal = wb!.Ranges(definedName);
        }

        foreach (IXLRange range in rangesLocal)
        {
            foreach (IXLRangeRow? row in range.Rows())
            {
                yield return row.Cells().Select(c => c.Value.ToString());
            }
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName) => throw new NotImplementedException();
}
