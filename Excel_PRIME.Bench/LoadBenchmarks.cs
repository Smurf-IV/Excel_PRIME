using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;


using Sylvan.Data.Excel;

using XlsxHelper;


namespace ExcelPRIME.Bench;
/*
| Method              | FileName             | Mean        | Error       | StdDev      | Ratio           | RatioSD | Gen0      | Gen1      | Gen2     | Allocated   | Alloc Ratio   |
|-------------------- |--------------------- |------------:|------------:|------------:|----------------:|--------:|----------:|----------:|---------:|------------:|--------------:|
| LoadWithSylvan      | Data/(...).xlsx [35] | 36,766.1 us | 25,222.5 us | 1,382.53 us |        baseline |         | 6357.1429 | 6071.4286 | 785.7143 | 45920.59 KB |               |
| LoadWithXlsxHelper  | Data/(...).xlsx [35] |    226.3 us |    112.2 us |     6.15 us | 162.576x faster |   6.50x |   12.6953 |   12.2070 |        - |    103.7 KB | 442.829x less |
| LoadWithExcel_Prime | Data/(...).xlsx [35] |  2,896.0 us |  5,381.3 us |   294.97 us |  12.779x faster |   1.14x |   42.9688 |   15.6250 |        - |   352.72 KB | 130.191x less |
| LoadWithFastExcel   | Data/(...).xlsx [35] |          NA |          NA |          NA |               ? |       ? |        NA |        NA |       NA |          NA |             ? |
|                     |                      |             |             |             |                 |         |           |           |          |             |               |
| LoadWithSylvan      | Data/(...).xlsx [39] | 12,959.8 us |  2,152.1 us |   117.97 us |        baseline |         | 1125.0000 | 1109.3750 |        - |  9352.07 KB |               |
| LoadWithXlsxHelper  | Data/(...).xlsx [39] |    179.3 us |    263.4 us |    14.44 us |   72.60x faster |   5.07x |    8.7891 |    8.3008 |        - |       75 KB | 124.688x less |
| LoadWithExcel_Prime | Data/(...).xlsx [39] |  2,618.8 us |  5,775.5 us |   316.58 us |    4.99x faster |   0.49x |   39.0625 |   15.6250 |        - |   323.96 KB |  28.868x less |
| LoadWithFastExcel   | Data/(...).xlsx [39] |          NA |          NA |          NA |               ? |       ? |        NA |        NA |       NA |          NA |             ? |
|                     |                      |             |             |             |                 |         |           |           |          |             |               |
| LoadWithSylvan      | Data/(...).xlsx [35] | 18,789.0 us |  7,430.1 us |   407.27 us |        baseline |         | 2468.7500 | 2406.2500 | 218.7500 | 18501.56 KB |               |
| LoadWithXlsxHelper  | Data/(...).xlsx [35] |    167.5 us |    258.4 us |    14.16 us | 112.731x faster |   8.41x |    9.2773 |    8.7891 |        - |    76.45 KB | 242.017x less |
| LoadWithExcel_Prime | Data/(...).xlsx [35] |  2,587.8 us |  4,751.0 us |   260.42 us |   7.307x faster |   0.62x |   39.0625 |   15.6250 |        - |   324.47 KB |  57.020x less |
| LoadWithFastExcel   | Data/(...).xlsx [35] |          NA |          NA |          NA |               ? |       ? |        NA |        NA |       NA |          NA |             ? |
 */

[ExcludeFromCodeCoverage]
public class LoadBenchmarks
{
    [Params(
        "Data/Blank Data 1 Million Rows.xlsx",
        "Data/sampledocs-50mb-xlsx-file.xlsx",
        "Data/sampledocs-50mb-xlsx-file-sst.xlsx"
        //,"Data/100mb.xlsx"
        )]
    public string FileName { get; set; }

    [Benchmark(Baseline = true)]
    public async Task LoadWithSylvan()
    {
        using var reader = await ExcelDataReader.CreateAsync(FileName);
    }

    [Benchmark]
    public void LoadWithXlsxHelper()
    {
        var reader = XlsxReader.OpenWorkbook(FileName);
    }

    [Benchmark]
    public async Task LoadWithExcel_Prime()
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(FileName);
    }

    //[Benchmark]   Does not work
    //public void LoadWithKapralExcel()
    //{
    //    using Kapral.FastExcel.FastExcel excel = new (FileName);
    //}

    [Benchmark]
    public void LoadWithFastExcel()
    {
        using FastExcel.FastExcel excel = new(new FileInfo(FileName));
    }

}
