using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;

using FastExcel;

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
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> AccessEveryCellSylvan()
    {
        int cells = 0;
        using ExcelDataReader reader = ExcelDataReader.Create(FileName);
        //using ExcelDataReader reader = await ExcelDataReader.CreateAsync(FileName).ConfigureAwait(true);
        do
        {
            while (await reader.ReadAsync().ConfigureAwait(true))
            {
                for (int ordinal = 0; ordinal < reader.RowFieldCount; ordinal++)
                {
                    _ = reader.GetExcelValue(ordinal);
                    cells++;
                }
            }
        } while (await reader.NextResultAsync().ConfigureAwait(true));
        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int AccessEveryCellXlsxHelper()
    {
        int cells = 0;
        Workbook workbook = XlsxReader.OpenWorkbook(FileName);
        foreach (XlsxHelper.Worksheet worksheet in workbook.Worksheets)
        {
            using WorksheetReader worksheetReader = worksheet.WorksheetReader;
            foreach (XlsxHelper.Row row in worksheetReader)
            {
                foreach (XlsxHelper.Cell cell in row.Cells)
                {
                    _ = cell.CellValue;
                    cells++;
                }
            }
        }

        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int AccessEveryCellFastExcel()
    {
        int cells = 0;
        string filePath = Path.Combine(Environment.CurrentDirectory, FileName);
        FileInfo inputFile = new FileInfo(filePath);

        using FastExcel.FastExcel excel = new(inputFile, true);

        foreach (FastExcel.Worksheet worksheet in excel.Worksheets)
        {
            worksheet.Read();
            foreach (FastExcel.Row row in worksheet.Rows)
            {
                foreach (FastExcel.Cell cell in row.Cells)
                {
                    _ = cell.Value;
                    cells++;
                }
            }
        }
        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> AccessEveryCellAsyncExcel_Prime()
    {
        int cells = 0;
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
                {
                    cells++;
                    if (cell != null)
                    {
                        // Because this returns upto the dimension of the sheet width
                        _ = cell.RawValue;
                    }
                }
                row.Dispose();
            }
        }
        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> AccessEveryCellExcel_Prime()
    {
        int cells = 0;
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
                {
                    cells++;
                    if (cell != null)
                    {
                        // Because this returns upto the dimension of the sheet width
                        _ = cell.RawValue;
                    }
                }
                row.Dispose();
            }
        }

        return cells;
    }
}

