using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Aspose.Cells;

using BenchmarkDotNet.Attributes;

using ExcelDataReader;

using ExcelPRIME;

namespace ExcelPRIMEXlsb.Bench;


[ExcludeFromCodeCoverage]
public class AccessEveryCellBenchmarksXlsb
{
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    public AccessEveryCellBenchmarksXlsb()
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    {
        // Needed to make the ExcelDataReaderWork !
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    private const string RootFolder = @"Data\";
    [Params(
        "Blank Data 1 Million Rows.xlsb",
        "sampledocs-50mb-xlsx-file.xlsb",
        "sampledocs-50mb-xlsx-file-sst.xlsb",
        "100mb.xlsb"
    )]
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    // ReSharper disable once PropertyCanBeMadeInitOnly.Global
    public string FileName { get; set; }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    [Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> SylvanRdr()
    {
        int cells = 0;
        // ReSharper disable once MethodHasAsyncOverload
        using Sylvan.Data.Excel.ExcelDataReader reader = Sylvan.Data.Excel.ExcelDataReader.Create(RootFolder + FileName);
        //using ExcelDataReader reader = await ExcelDataReader.CreateAsync(RootFolder + FileName).ConfigureAwait(true);
        do
        {
            while (await reader.ReadAsync().ConfigureAwait(true))
            {
                for (int ordinal = 0; ordinal < reader.RowFieldCount; ordinal++)
                {
                    string? value = reader.GetExcelValue(ordinal).ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        cells++;
                    }
                }
            }
        } while (await reader.NextResultAsync().ConfigureAwait(true));

        return cells;
    }


    //[Benchmark]
    /*
     * Benchmarks with issues:
       AccessEveryCellBenchmarksXlsb.AccessEveryCellExcelDataReader: ShortRun(IterationCount=3, LaunchCount=1, WarmupCount=3) [FileName=Blank(...).xlsb [30]]
       AccessEveryCellBenchmarksXlsb.AccessEveryCellExcelDataReader: ShortRun(IterationCount=3, LaunchCount=1, WarmupCount=3) [FileName=sampl(...).xlsb [34]]
       AccessEveryCellBenchmarksXlsb.AccessEveryCellExcelDataReader: ShortRun(IterationCount=3, LaunchCount=1, WarmupCount=3) [FileName=sampl(...).xlsb [30]]
     */
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int ExcelDataReader()
    {
        int cells = 0;
        using FileStream stream = File.Open(RootFolder + FileName, FileMode.Open, FileAccess.Read);
        // Auto-detect format, supports:
        //  - Binary Excel files (2.0-2003 format; *.xls)
        //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
        using IExcelDataReader? reader = ExcelReaderFactory.CreateReader(stream);
        while (reader.Read())
        {
            int cols = reader.FieldCount;
            for (int c = 0; c < cols; c++)
            {
                object cell = reader.GetValue(c);
                string? value = cell.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    cells++;
                }
            }
        }

        return cells;
    }


    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int Aspose()
    {
        int cells = 0;
        using Workbook wb = new Workbook(RootFolder + FileName);
        foreach (Worksheet? ws in wb.Worksheets)
        {
            Cells? wsCells = ws.Cells;

            int maxRow = wsCells.MaxDataRow;      // highest row index with data
            int maxCol = wsCells.MaxDataColumn;   // highest column index with data

            for (int r = 0; r <= maxRow; r++)
            {
                // Optionally skip fully empty rows
                bool rowHasData = false;
                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = wsCells[r, c];
                    if (cell.Value != null)
                    {
                        rowHasData = true;
                        break;
                    }
                }
                if (!rowHasData)
                    continue;

                // Process row r
                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = wsCells[r, c];
                    string? value = cell.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        cells++;
                    }
                }
            }
        }
        return cells;
    }

    [Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> AsyncExcel_PrimeXlsb()
    {
        int cells = 0;
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(RootFolder + FileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName);
            await foreach (IRowAsync? row in worksheet!.GetRowDataAsync())
            {
                if (row == null)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                if (rowCells != null)
                {
                    foreach (ICell? cell in rowCells)
                    {
                        // Because this returns upto the dimension of the sheet width
                        string? value = cell?.CellValue.ToString();
                        if (!string.IsNullOrEmpty(value))
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
    public int Excel_PrimeXlsb()
    {
        int cells = 0;
        using Excel_PRIMEXlsb workbook = new();
        workbook.Open(RootFolder + FileName);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = workbook.GetSheet(sheetName);
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row == null)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                IReadOnlyList<ICell?>? rowCells = row.GetAllCells();
                if (rowCells != null)
                {
                    foreach (ICell? cell in rowCells)
                    {
                        // Because this returns upto the dimension of the sheet width
                        string? value = cell?.CellValue.ToString();
                        if (!string.IsNullOrEmpty(value))
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
    // Between 5 -> 10% slower than running through in ForwardOnlyMode*2.
    // Not bad considering it is using the HDD for the passes ;-)
    // BUT:  100mb.xlsx = `2.65x slower`;  Compared to `1.60x slower` for ForwardOnlyMode*1
    // Memory is between 80% and 110% more
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> PrlAsyncExcel_PrimeXlsbTwice()
    {
        using Excel_PRIMEXlsb workbook = new();
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
                    if (row == null)
                    {
                        // Because this returns upto the dimension of the sheet Height
                        break;
                    }

                    IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                    if (rowCells != null)
                    {
                        foreach (ICell? cell in rowCells)
                        {
                            // Because this returns upto the dimension of the sheet width
                            string? value = cell?.CellValue.ToString();
                            if (!string.IsNullOrEmpty(value))
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
                    if (row == null)
                    {
                        // Because this returns upto the dimension of the sheet Height
                        break;
                    }

                    IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                    if (rowCells != null)
                    {
                        foreach (ICell? cell in rowCells)
                        {
                            // Because this returns upto the dimension of the sheet width
                            string? value = cell?.CellValue.ToString();
                            if (!string.IsNullOrEmpty(value))
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
