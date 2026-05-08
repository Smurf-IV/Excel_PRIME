using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;

using Sylvan.Data.Excel;

using XlsxHelper;

namespace ExcelPRIME.Bench;


[ExcludeFromCodeCoverage]
public class AccessEveryCellBenchmarks
{
    private const string RootFolder = @"Data\";
    [Params(
        "Blank Data 1 Million Rows.xlsx",
        "sampledocs-50mb-xlsx-file.xlsx",
        "sampledocs-50mb-xlsx-file-sst.xlsx",
        "100mb.xlsx"
    )]
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    // ReSharper disable once PropertyCanBeMadeInitOnly.Global
    public string FileName { get; set; }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    
    static ExcelDataReaderOptions ErrorsAsString = 
        new ExcelDataReaderOptions { 
            // formula errors read as their string values
            FormulaErrorHandling = FormulaErrorHandling.String 
        };

    [Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> SylvanRdr()
    {
        int cells = 0;
        using ExcelDataReader reader = await ExcelDataReader.CreateAsync(RootFolder + FileName, ErrorsAsString).ConfigureAwait(true);
        do
        {
            // include the headers in the populated cell count
            cells += reader.FieldCount;
            while (await reader.ReadAsync().ConfigureAwait(true))
            {
                for (int ordinal = 0; ordinal < reader.RowFieldCount; ordinal++)
                {
                    // this also passes the unit test, and is much more efficient
                    //if (!reader.IsDBNull(ordinal)) 
                    if (!string.IsNullOrEmpty(reader.GetString(ordinal)))
                    {
                        cells++;
                    }
                }
            }
        } while (await reader.NextResultAsync().ConfigureAwait(true));

        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int XlsxHelper()
    {
        int cells = 0;
        using Workbook workbook = XlsxReader.OpenWorkbook(RootFolder + FileName);
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            using WorksheetReader worksheetReader = worksheet.WorksheetReader;
            foreach (Row row in worksheetReader)
            {
                foreach (Cell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.CellValue))
                    {
                        cells++;
                    }
                }
            }
        }

        return cells;
    }

    //[Benchmark]
    //[MethodImpl(MethodImplOptions.NoOptimization)]
    //public int FastExcel()
    //{
    //    int cells = 0;
    //    string filePath = Path.Combine(Environment.CurrentDirectory, RootFolder + FileName);
    //    FileInfo inputFile = new FileInfo(filePath);

    //    using FastExcel.FastExcel excel = new(inputFile, true);

    //    foreach (FastExcel.Worksheet worksheet in excel.Worksheets)
    //    {
    //        worksheet.Read();
    //        foreach (FastExcel.Row row in worksheet.Rows)
    //        {
    //            foreach (FastExcel.Cell cell in row.Cells)
    //            {
    //                _ = cell;
    //                cells++;
    //            }
    //        }
    //    }
    //    return cells;
    //}

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> AsyncExcel_Prime()
    {
        int cells = 0;
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(RootFolder + FileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName);
            await foreach (IRowAsync? row in worksheet!.GetRowDataAsync())
            {
                if (row is null or INullRowAsync)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                if (rowCells != null)
                {
                    foreach (ICell? cell in rowCells)
                    {
                        // Because this returns upto the dimension of the sheet width
                        if (!string.IsNullOrEmpty(cell?.CellValue?.ToString()))
                        {
                            cells++;
                        }
                    }
                }

                row.Dispose();
            }
        }
        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int Excel_Prime()
    {
        int cells = 0;
        using Excel_PRIME workbook = new();
        workbook.Open(RootFolder + FileName);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = workbook.GetSheet(sheetName);
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row is null or INullRow)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                IReadOnlyList<ICell?>? rowCells = row.GetAllCells();
                if (rowCells != null)
                {
                    cells += rowCells.Count(cell => !string.IsNullOrEmpty(cell?.CellValue?.ToString()));
                }
                row.Dispose();
            }
        }

        return cells;
    }

    [Benchmark]
    // Between 5 -> 10% slower than running through in ForwardOnlyMode*2.
    // Not bad considering it is using the HDD for the passes ;-)
    // BUT:  100mb.xlsx = `2.65x slower`;  Compared to `1.60x slower` for ForwardOnlyMode*1
    // Memory is between 80% and 110% more
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> PrlAsyncExcel_PrimeTwice()
    {
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(RootFolder + FileName, options: new Options { AccessExcelFileInForwardOnlyMode = false }).ConfigureAwait(true);

        ParallelOptions parallelOptions = new()
        {
            MaxDegreeOfParallelism = Environment.ProcessorCount - 1,
            CancellationToken = CancellationToken.None
        };
        int cells = 0;
        await Parallel.ForEachAsync(workbook.SheetNames(),
            parallelOptions,
            async (sheetName, ct) =>
            {
                using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName, false, ct);
                await foreach (IRowAsync? row in worksheet!.GetRowDataAsync(ct: ct))
                {
                    if (row is null or INullRowAsync)
                    {
                        // Because this returns upto the dimension of the sheet Height
                        break;
                    }

                    IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync(ct).ConfigureAwait(true);
                    if (rowCells != null)
                    {
                        foreach (ICell? cell in rowCells)
                        {
                            // Because this returns upto the dimension of the sheet width
                            if (!string.IsNullOrEmpty(cell?.CellValue?.ToString()))
                            {
                                Interlocked.Increment(ref cells);
                            }
                        }
                    }

                    row.Dispose();
                }
            });

        cells = 0;
        await Parallel.ForEachAsync(workbook.SheetNames(),
            parallelOptions,
            async (sheetName, ct) =>
            {
                using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName, false, ct);
                await foreach (IRowAsync? row in worksheet!.GetRowDataAsync(ct: ct))
                {
                    if (row is null or INullRowAsync)
                    {
                        // Because this returns upto the dimension of the sheet Height
                        break;
                    }

                    IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync(ct).ConfigureAwait(true);
                    if (rowCells != null)
                    {
                        foreach (ICell? cell in rowCells)
                        {
                            // Because this returns upto the dimension of the sheet width
                            if (!string.IsNullOrEmpty(cell?.CellValue?.ToString()))
                            {
                                Interlocked.Increment(ref cells);
                            }
                        }
                    }

                    row.Dispose();
                }
            });
        return cells;
    }
}
