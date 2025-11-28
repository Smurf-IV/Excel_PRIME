using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.VisualStudio.OLE.Interop;

using Spire.Xls;
using Spire.Xls.Core;

namespace ExcelPRIME.RangeBench;

public class GetRangeFreeSpire : IGetRange
{
    private readonly Workbook book = new Workbook();
    public void Dispose() => book.Dispose();

    public bool LoadFile(string fullPath)
    {
        book.LoadFromFile(fullPath);
        return true;
    }

    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, int? localSheetId = null)
    {
        INamedRange namedRange;
        if ( localSheetId.HasValue)
        {
            //Get a specific named range in the worksheet
            Worksheet sheet = book.Worksheets[localSheetId.Value];
            namedRange = sheet.Names.GetByName(definedName);
        }
        else
        {
            //Get a specific named range in the workbook
            namedRange = book.NameRanges[definedName];
        }
        if (namedRange != null)
        {
            //Get range
            foreach (IXLSRange row in namedRange.RefersToRange.Rows)
            {
                foreach (CellRange cells in row.CellList)
                {
                    yield return cells.Select(c => c.Value.ToString());
                }
            }
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName) => throw new NotImplementedException();
}
