using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

using Microsoft.VSDiagnostics;

using Sylvan.Data.Excel;

using XlsxHelper;

namespace ExcelPRIME.Bench;

[ExcludeFromCodeCoverage]
public class AccessEveryCellBenchmarks
{
    [Params(
        "Data/Blank Data 1 Million Rows.xlsx",
        "Data/sampledocs-50mb-xlsx-file.xlsx",
        "Data/sampledocs-50mb-xlsx-file-sst.xlsx",
        "Data/100mb.xlsx"
    )]
    public string FileName { get; set; }

    [Benchmark(Baseline = true)]
    public async Task AccessEveryCellSylvan()
    {
        using ExcelDataReader reader = await ExcelDataReader.CreateAsync(FileName).ConfigureAwait(true);
        do
        {
            while (await reader.ReadAsync().ConfigureAwait(true))
            {
                if (reader.RowFieldCount > 0)
                {
                    _ = reader.GetExcelValue(1);
                }
            }
        } while (await reader.NextResultAsync().ConfigureAwait(true));
    }

    [Benchmark]
    public void AccessEveryCellXlsxHelper()
    {
        Workbook workbook = XlsxReader.OpenWorkbook(FileName);
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            using WorksheetReader worksheetReader = worksheet.WorksheetReader;
            foreach (Row row in worksheetReader)
            {
                if (row.Cells.Any())
                {
                    _ = row.Cells[0].CellValue;
                }
            }
        }
    }

    [Benchmark]
    public async Task AccessEveryCellExcel_Prime()
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(FileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(sheetName);
            await foreach (IRow? row in worksheet!.GetRowDataAsync())
            {
                if (row == null)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                ICell? cell = await row.GetCellAsync(1).ConfigureAwait(true);
                if (cell != null)
                {   // Because this returns upto the dimension of the sheet width
                    _ = cell.RawValue!.ToString();
                }
            }
        }
    }
}

