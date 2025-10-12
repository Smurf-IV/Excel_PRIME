using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;

using Microsoft.VSDiagnostics;

using Sylvan.Data.Excel;

using XlsxHelper;

namespace ExcelPRIME.Bench;
/*
| Method              | FileName             | Mean           | Error        | StdDev       | Median         | Ratio | RatioSD | Gen0        | Gen1       | Gen2      | Allocated   | Alloc Ratio |
|-------------------- |--------------------- |---------------:|-------------:|-------------:|---------------:|------:|--------:|------------:|-----------:|----------:|------------:|------------:|
| LoadWithExcel_Prime | Data/100mb.xlsx      | 2,652,606.9 us | 52,991.58 us | 52,044.82 us | 2,652,595.6 us | 52.89 |    2.42 | 101000.0000 | 71000.0000 | 4000.0000 | 829111.1 KB |      11.681 |
| LoadWithSylvan      | Data/100mb.xlsx      |    50,235.3 us |    994.60 us |  2,054.03 us |    50,313.7 us |  1.00 |    0.06 |   9900.0000 |  9800.0000 | 1300.0000 | 70977.77 KB |       1.000 |
| LoadWithXlsxHelper  | Data/100mb.xlsx      |       808.4 us |     29.91 us |     88.18 us |       757.6 us |  0.02 |    0.00 |     21.4844 |    19.5313 |         - |   183.64 KB |       0.003 |
|                     |                      |                |              |              |                |       |         |             |            |           |             |             |
| LoadWithSylvan      | Data/(...).xlsx [35] |    39,568.3 us |    781.67 us |  1,764.36 us |    39,723.0 us | 1.002 |    0.06 |   6384.6154 |  6076.9231 |  846.1538 | 45921.85 KB |       1.000 |
| LoadWithExcel_Prime | Data/(...).xlsx [35] |     5,591.5 us |     82.44 us |     84.66 us |     5,606.8 us | 0.142 |    0.01 |    218.7500 |   203.1250 |         - |  1819.39 KB |       0.040 |
| LoadWithXlsxHelper  | Data/(...).xlsx [35] |       228.9 us |      4.38 us |      5.70 us |       227.5 us | 0.006 |    0.00 |     12.2070 |    11.7188 |         - |   103.54 KB |       0.002 |
|                     |                      |                |              |              |                |       |         |             |            |           |             |             |
| LoadWithSylvan      | Data/(...).xlsx [39] |    12,606.5 us |    236.85 us |    253.43 us |    12,574.4 us |  1.00 |    0.03 |   1125.0000 |  1109.3750 |         - |     9352 KB |       1.000 |
| LoadWithExcel_Prime | Data/(...).xlsx [39] |     2,960.7 us |     76.58 us |    222.18 us |     2,855.0 us |  0.23 |    0.02 |     54.6875 |    23.4375 |         - |   476.97 KB |       0.051 |
| LoadWithXlsxHelper  | Data/(...).xlsx [39] |       161.6 us |      3.15 us |      3.10 us |       161.2 us |  0.01 |    0.00 |      8.7891 |     8.3008 |         - |    74.97 KB |       0.008 |
|                     |                      |                |              |              |                |       |         |             |            |           |             |             |
| LoadWithSylvan      | Data/(...).xlsx [35] |    18,688.2 us |    340.12 us |    746.56 us |    18,586.3 us | 1.002 |    0.06 |   2468.7500 |  2375.0000 |  218.7500 | 18501.64 KB |       1.000 |
| LoadWithExcel_Prime | Data/(...).xlsx [35] |     2,980.0 us |     85.21 us |    251.26 us |     2,858.7 us | 0.160 |    0.01 |     50.7813 |    15.6250 |         - |   430.99 KB |       0.023 |
| LoadWithXlsxHelper  | Data/(...).xlsx [35] |       170.5 us |      2.98 us |      4.27 us |       170.0 us | 0.009 |    0.00 |      9.2773 |     8.7891 |         - |    76.45 KB |       0.004 |
 */

[ExcludeFromCodeCoverage]
public class LoadBenchmarks
{
    [Params(
        "Data/Blank Data 1 Million Rows.xlsx",
        "Data/sampledocs-50mb-xlsx-file.xlsx",
        "Data/sampledocs-50mb-xlsx-file-sst.xlsx",
        "Data/100mb.xlsx"
        )]
    public string FileName { get; set; }

    //[Benchmark(Baseline = true)]
    public async Task LoadWithSylvan()
    {
        using var reader = await ExcelDataReader.CreateAsync(FileName);
    }

    //[Benchmark]
    public void LoadWithXlsxHelper()
    {
        var reader = XlsxReader.OpenWorkbook(FileName);
    }

    //[Benchmark]
    public async Task LoadWithExcel_Prime()
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(FileName);
    }
}
