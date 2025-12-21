using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

using BenchmarkDotNet.Attributes;

using Microsoft.VSDiagnostics;

namespace ExcelPRIME.Bench;

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor

[ExcludeFromCodeCoverage]
[SimpleJob(warmupCount: 1, iterationCount: 3)]
[CPUUsageDiagnoser]
[MemoryDiagnoser]
public class StringAllocationBenchmark
{
    private const string RootFolder = @"Data\";

    [Params("100mb.xlsx")]
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
    // ReSharper disable once PropertyCanBeMadeInitOnly.Global
    public string FileName { get; set; }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.

    //[Benchmark(Baseline = true)]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int AccessStringCellsSequential()
    {
        int totalChars = 0;
        int stringCellCount = 0;

        using Excel_PRIME workbook = new();
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

                foreach (ICell? cell in row.GetAllCells())
                {
                    if (cell is { RawExcelType: CellType.InlineString, RawValue: string s })
                    {
                        totalChars += s.Length;
                        stringCellCount++;
                    }
                }

                row.Dispose();
                if (stringCellCount > 5000)
                {
                    break;
                }
            }

            if (stringCellCount > 5000)
            {
                break;
            }
        }

        return totalChars;
    }

    //[Benchmark]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int AccessStringCellsWithToString()
    {
        int totalChars = 0;
        int stringCellCount = 0;

        using Excel_PRIME workbook = new();
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

                foreach (ICell? cell in row.GetAllCells())
                {
                    if (cell != null)
                    {
                        string? cellStr = cell.RawValue?.ToString();
                        if (!string.IsNullOrEmpty(cellStr))
                        {
                            totalChars += cellStr.Length;
                            stringCellCount++;
                        }
                    }
                }

                row.Dispose();
                if (stringCellCount > 5000)
                {
                    break;
                }
            }

            if (stringCellCount > 5000)
            {
                break;
            }
        }

        return totalChars;
    }
}

#pragma warning restore CS8618
