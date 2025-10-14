using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;

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
                for (int ordinal = 0; ordinal < reader.RowFieldCount; ordinal++)
                {
                    _ = reader.GetExcelValue(ordinal);
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
                foreach (Cell cell in row.Cells)
                {
                    _ = cell.CellValue;
                }
            }
        }
    }

    [Benchmark]
    public async Task AccessEveryCellAsyncExcel_Prime()
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

                await foreach (ICell? cell in row.GetAllCellsAsync())
                {   // Because this returns upto the dimension of the sheet width
                    _ = cell?.RawValue!.ToString();
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
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row == null)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                foreach (ICell? cell in row.GetAllCells())
                {   // Because this returns upto the dimension of the sheet width
                    _ = cell?.RawValue!.ToString();
                }
            }
        }
    }
}

