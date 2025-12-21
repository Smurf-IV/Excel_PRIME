using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;

using ExcelPRIME;

using Microsoft.VSDiagnostics;

namespace ExcelPRIME.Bench;
/// <summary>
/// Profile XLSB non-numeric bottlenecks:
/// - String extraction and encoding (Encoding.Unicode.GetString)
/// - Shared string lookups and lazy loading
/// - Column name generation and caching
/// - RawValue ToString() conversions
/// </summary>
[ExcludeFromCodeCoverage]
[SimpleJob(warmupCount: 2, iterationCount: 3)]
[CPUUsageDiagnoser]
public class XlsbBottleneckProfileBenchmark
{
    private const string RootFolder = @"Data\";
    [Params("100mb.xlsb")]
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    // ReSharper disable once PropertyCanBeMadeInitOnly.Global
    public string FileName { get; set; }

#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    /// <summary>
    /// Baseline async: Access all cell properties (RawValue.ToString)
    /// </summary>
    //[Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public async Task<int> AccessEveryCellXlsb_Prime_Async()
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
                {
                    break;
                }

                await foreach (ICell? cell in row.GetAllCellsAsync())
                {
                    if (!string.IsNullOrEmpty(cell?.RawValue?.ToString()))
                    {
                        cells++;
                    }
                }

                row.Dispose();
            }
        }

        return cells;
    }

    /// <summary>
    /// Synchronous variant of baseline for comparison
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int AccessEveryCellXlsb_Prime_Sync()
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
                {
                    break;
                }

                cells += row.GetAllCells().Count(cell => !string.IsNullOrEmpty(cell?.RawValue?.ToString()));
                row.Dispose();
            }
        }

        return cells;
    }

    /// <summary>
    /// Profile RawValue access only (no ToString conversions)
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int ProfileXlsbAccessCellValuesOnly()
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
                {
                    break;
                }

                cells += row.GetAllCells().Count(cell => cell?.RawValue != null);
                row.Dispose();
            }
        }

        return cells;
    }

    /// <summary>
    /// Profile column name generation and caching
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int ProfileXlsbAccessColumnLettersOnly()
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
                {
                    break;
                }

                cells += row.GetAllCells().Count(cell => cell?.ColumnLetters != null);
                row.Dispose();
            }
        }

        return cells;
    }

    /// <summary>
    /// Profile ToString conversions (most expensive operation)
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int ProfileXlsbStringConversions()
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
                {
                    break;
                }

                cells += row.GetAllCells().Count(cell => !string.IsNullOrEmpty(cell?.RawValue?.ToString()));
                row.Dispose();
            }
        }

        return cells;
    }

    /// <summary>
    /// Profile with CellConversion.Number enabled
    /// </summary>
    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int ProfileXlsbWithNumberConversion()
    {
        int cells = 0;
        using Excel_PRIMEXlsb workbook = new();
        workbook.Open(RootFolder + FileName, new Options { CellConversionType = CellConversion.Number });
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = workbook.GetSheet(sheetName);
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row == null)
                {
                    break;
                }

                cells += row.GetAllCells().Count(cell => cell?.RawValue != null);
                row.Dispose();
            }
        }

        return cells;
    }
}