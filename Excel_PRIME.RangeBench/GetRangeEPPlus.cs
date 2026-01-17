using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using OfficeOpenXml;

namespace ExcelPRIME.RangeBench;

public class GREPPlus : IGetRange
{
    private ExcelPackage? excelPackage;

    public void Dispose()
    {
        excelPackage?.Dispose();
        excelPackage = null;
    }

    public bool LoadFile(string fullPath)
    {
        excelPackage = new ExcelPackage(new FileInfo(fullPath));
        return excelPackage != null;
    }

    public IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, string? sheetName = null)
    {
        ExcelNamedRange? namedRange;
        // worksheet scope
        if (sheetName != null)
        {
            ExcelWorksheet? worksheet = excelPackage!.Workbook.Worksheets[sheetName];
            namedRange = worksheet.Names[definedName];
        }
        else
        {
            // workbook scope
            namedRange = excelPackage!.Workbook.Names[definedName];
        }
        if (namedRange != null)
        {
            ExcelNamedRange range = namedRange;
            for (int row = range.Start.Row; row <= range.End.Row; row++)
            {
                List<object?> rowData = new List<object?>(range.End.Column - range.Start.Column);
                for (int col = range.Start.Column; col <= range.End.Column; col++)
                {
                    rowData.Add(range.Worksheet.Cells[row, col].Value.ToString());
                }
                yield return rowData;
            }
        }
    }

    public IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName) => throw new NotImplementedException();
}
