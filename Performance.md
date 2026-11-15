<!-- TOC-->
- [2026-10-18 - V5 nuget update](#2026-10-18---v5-nuget-update)
- [2026-11-12 - V6 Fast Alpha](#2026-11-12---v6-fast-alpha)
- [2026-11-13 - V6 - Excel_FAST - Beta](#2026-11-13---v6---excel_fast---beta)
<!-- TOC -->



# 2026-10-18 - V5 nuget update
- Update nugets
- Made benchmarks run .net**9** and .net**10** as well
- **Memory & Allocation Optimizations**
- Some tinkering in the way memory is used and allocated, especially in the Xlsb workflow
- Eliminated String Caching Overhead: Removed the ConcurrentDictionary cache for small strings in CellValue, which introduced locking and hashing overhead during high-volume parsing.
- Simplified SIMD Decoding: Reverted SimdStringDecoder.cs to use Vector<ushort> instead of explicit Vector512/Vector256/Vector128 branching, reducing overhead for common short strings.
- Optimized Column Parsing: Replaced SearchValues with simple loops in ExcelColumns.cs for cell reference parsing, as the overhead of SearchValues was counterproductive for 1-3 character strings.
- Tuned Method Inlining: Removed AggressiveOptimization from several complex methods to allow better JIT tiered compilation and reduce instruction cache pressure.
```
Job=net10.0  Runtime=.NET 10.0  IterationCount=3
LaunchCount=1  WarmupCount=3

| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  21000.0000 | 20000.0000 | 2000.0000 |  373.63 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.59x slower | 188000.0000 |  1000.0000 |         - | 3380.58 MB |  9.05x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.45x slower |  44000.0000 | 37000.0000 | 4000.0000 |  729.87 MB |  1.95x more |
| Excel_Prime              | 100mb.xlsx           | 1.38x slower |  43000.0000 | 36000.0000 | 4000.0000 |  726.88 MB |  1.95x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.34x slower |  51000.0000 | 39000.0000 | 4000.0000 |  869.05 MB |  2.33x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  12000.0000 |  1000.0000 |         - |  221.89 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.01x faster |  96000.0000 |  1000.0000 |         - | 1739.24 MB |  7.84x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.15x slower |  16000.0000 |  1000.0000 |         - |  297.57 MB |  1.34x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.03x slower |  16000.0000 |  1000.0000 |         - |  293.16 MB |  1.32x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.63x slower |  34000.0000 |  1000.0000 |         - |  624.92 MB |  2.82x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |   9000.0000 |          - |         - |     178 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.09x faster |  44000.0000 |          - |         - |  799.73 MB |  4.49x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.09x slower |  12000.0000 |  1000.0000 |         - |   215.6 MB |  1.21x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.06x faster |  11000.0000 |  1000.0000 |         - |  212.74 MB |  1.20x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.32x slower |  24000.0000 |  1000.0000 |         - |  442.97 MB |  2.49x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |   9000.0000 |          - |         - |  175.98 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.19x faster |  41000.0000 |          - |         - |  742.13 MB |  4.22x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.09x slower |  14000.0000 |  1000.0000 |         - |  253.73 MB |  1.44x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.07x faster |  14000.0000 |          - |         - |  251.66 MB |  1.43x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.29x slower |  28000.0000 |  1000.0000 |         - |  515.56 MB |  2.93x more |
```
```
|----------------- |---------:|--------:|------:|------------:|------------:|
| Baseline         | 103.4 ms | 1.49 ms |  1.00 |   246.55 KB |        1.00 |
| SylvanXlsx       | 155.3 ms | 2.17 ms |  1.50 |   667.01 KB |        2.71 |
| SylvanXlsx_BindT | 162.3 ms | 3.00 ms |  1.57 | 11932.66 KB |       48.40 |
| SylvanXlsxObj    | 167.5 ms | 1.19 ms |  1.62 | 14475.92 KB |       58.71 |
| XlsxHelperXlsx   | 200.2 ms | 1.77 ms |  1.94 | 93307.68 KB |      378.45 |
| PrimeXlsxType    | 200.2 ms | 1.91 ms |  1.94 |  2755.06 KB |       11.17 |
| PrimeXlsxObj     | 205.8 ms | 2.12 ms |  1.99 | 16970.21 KB |       68.83 |
| PrimeXlsx        | 212.8 ms | 3.36 ms |  2.06 | 25075.34 KB |      101.70 |
```
```
| Method                       | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|----------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                    | 100mb.xlsb           |     baseline |  19000.0000 | 18000.0000 | 4000.0000 |  335.55 MB |             |
| AsyncExcel_PrimeXlsb         | 100mb.xlsb           | 2.47x slower | 211000.0000 | 46000.0000 | 6000.0000 | 3693.66 MB | 11.01x more |
| Excel_PrimeXlsb              | 100mb.xlsb           | 1.80x slower |  95000.0000 | 39000.0000 | 2000.0000 | 1694.93 MB |  5.05x more |
| PrlAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           | 3.68x slower | 374000.0000 | 50000.0000 | 6000.0000 | 6611.71 MB | 19.70x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | Blank(...).xlsb [30] |     baseline |  16000.0000 |  1000.0000 |         - |  301.86 MB |             |
| AsyncExcel_PrimeXlsb         | Blank(...).xlsb [30] | 3.10x slower | 212000.0000 |  1000.0000 |         - |  3805.1 MB | 12.61x more |
| Excel_PrimeXlsb              | Blank(...).xlsb [30] | 1.96x slower |  77000.0000 |  1000.0000 |         - | 1396.66 MB |  4.63x more |
| PrlAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] | 6.40x slower | 424000.0000 |  1000.0000 |         - | 7612.54 MB | 25.22x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [34] |     baseline |  14000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [34] | 2.47x slower | 115000.0000 |  1000.0000 |         - |  2078.6 MB |  7.93x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [34] | 1.67x slower |  43000.0000 |          - |         - |  773.55 MB |  2.95x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] | 4.95x slower | 231000.0000 |  1000.0000 |         - | 4157.75 MB | 15.86x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [30] |     baseline |  14000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [30] | 2.48x slower | 115000.0000 |  1000.0000 |         - | 2078.58 MB |  7.93x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [30] | 1.65x slower |  43000.0000 |          - |         - |  773.55 MB |  2.95x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] | 4.89x slower | 231000.0000 |  1000.0000 |         - | 4157.75 MB | 15.86x more |
```
```
| Method      | ranger           | Mean    | Error    | StdDev   | Gen0        | Gen1        | Gen2      | Allocated |
|------------ |----------------- |--------:|---------:|---------:|------------:|------------:|----------:|----------:|
| Access100mb | GRAsposeCells    | 3.866 s | 1.1969 s | 0.0656 s | 113000.0000 | 113000.0000 | 8000.0000 |   2.03 GB |
| Access100mb | GRFreeSpire      | 4.964 s | 0.3388 s | 0.0186 s | 155000.0000 | 154000.0000 | 7000.0000 |   3.72 GB |
| Access100mb | GRExcelPrimeXlsb | 1.680 s | 0.3447 s | 0.0189 s | 121000.0000 |  38000.0000 | 5000.0000 |   2.05 GB |
```

# 2026-11-12 - V6 Fast Alpha
- Restructure the construction of the Sync and Async base classes to reduce overhead and improve performance.
- Update Benchmark nugets
- Introduced ExcelFAST zero-allocation pooled streaming subsystem for .xlsx and .xlsb formats.
- Exposes forward-only sequential streaming interfaces `IExcelFast`, `ISheetFast`, and `IRowFast`.
- Utilizes `ArrayPool<Cell>.Shared` and thread-static buffers for flat memory usage during high-throughput sheet scanning.
- Delivers zero per-row GC allocation during sequential cell consumption using `PooledCellView` and rented array segments.
- Use `ReadOnlySpan<Cell>` instead of `ArraySegment<Cell>`
    - Using `foreach` over `ReadOnlySpan<Cell>` is guaranteed to be allocation-free and boxing-free.
    - If you need LINQ-style processing on a span without allocations, iterate using `for` or `foreach` directly over the span

```
Job=net10.0  Runtime=.NET 10.0  IterationCount=3
LaunchCount=1  WarmupCount=3

| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  21000.0000 | 20000.0000 | 2000.0000 |  373.63 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.55x slower | 188000.0000 |  1000.0000 |         - | 3380.58 MB |  9.05x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.44x slower |  44000.0000 | 37000.0000 | 4000.0000 |  729.87 MB |  1.95x more |
| Excel_Prime              | 100mb.xlsx           | 1.29x slower |  43000.0000 | 36000.0000 | 4000.0000 |  726.88 MB |  1.95x more |
| ExcelFastXlsx            | 100mb.xlsx           | 1.13x slower |  45000.0000 | 38000.0000 | 5000.0000 |  730.93 MB |  1.96x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.36x slower |  52000.0000 | 40000.0000 | 5000.0000 |  869.14 MB |  2.33x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  12000.0000 |  1000.0000 |         - |  221.89 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.00x faster |  96000.0000 |  1000.0000 |         - | 1739.24 MB |  7.84x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.19x slower |  16000.0000 |  9000.0000 | 1000.0000 |  297.02 MB |  1.34x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.02x slower |  16000.0000 |  1000.0000 |         - |  293.16 MB |  1.32x more |
| ExcelFastXlsx            | Blank(...).xlsx [30] | 1.19x faster |  16000.0000 |  1000.0000 |         - |  293.66 MB |  1.32x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.67x slower |  34000.0000 |  1000.0000 |         - |  624.71 MB |  2.82x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |   9000.0000 |          - |         - |  177.99 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.11x faster |  44000.0000 |          - |         - |  799.73 MB |  4.49x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.09x slower |  12000.0000 |  1000.0000 |         - |  215.59 MB |  1.21x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.06x faster |  11000.0000 |  1000.0000 |         - |  212.74 MB |  1.20x more |
| ExcelFastXlsx            | sampl(...).xlsx [34] | 1.30x faster |  11000.0000 |  1000.0000 |         - |  212.74 MB |  1.20x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.34x slower |  24000.0000 |  1000.0000 |         - |  442.85 MB |  2.49x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |   9000.0000 |          - |         - |  175.97 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.05x faster |  41000.0000 |          - |         - |  742.13 MB |  4.22x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.23x slower |  14000.0000 |  1000.0000 |         - |  253.73 MB |  1.44x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.03x slower |  14000.0000 |          - |         - |  251.66 MB |  1.43x more |
| ExcelFastXlsx            | sampl(...).xlsx [30] | 1.11x faster |  14000.0000 |          - |         - |  251.66 MB |  1.43x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.53x slower |  28000.0000 |  1000.0000 |         - |  515.46 MB |  2.93x more |
```
```
| Method                       | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|----------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                    | 100mb.xlsb           |     baseline |  19000.0000 | 18000.0000 | 4000.0000 |  335.54 MB |             |
| AsyncExcel_PrimeXlsb         | 100mb.xlsb           | 2.51x slower | 211000.0000 | 46000.0000 | 6000.0000 | 3693.67 MB | 11.01x more |
| Excel_PrimeXlsb              | 100mb.xlsb           | 1.69x slower |  91000.0000 | 39000.0000 | 2000.0000 | 1618.23 MB |  4.82x more |
| ExcelFastXlsb                | 100mb.xlsb           | 1.74x slower |  67000.0000 | 45000.0000 | 2000.0000 | 1190.82 MB |  3.55x more |
| PrlAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           | 3.73x slower | 373000.0000 | 41000.0000 | 5000.0000 | 6611.69 MB | 19.70x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | Blank(...).xlsb [30] |     baseline |  16000.0000 |  1000.0000 |         - |  301.86 MB |             |
| AsyncExcel_PrimeXlsb         | Blank(...).xlsb [30] | 3.19x slower | 212000.0000 |  1000.0000 |         - | 3805.08 MB | 12.61x more |
| Excel_PrimeXlsb              | Blank(...).xlsb [30] | 1.85x slower |  73000.0000 |  1000.0000 |         - |  1326.5 MB |  4.39x more |
| ExcelFastXlsb                | Blank(...).xlsb [30] | 1.93x slower |  38000.0000 |  1000.0000 |         - |  686.49 MB |  2.27x more |
| PrlAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] | 6.65x slower | 424000.0000 |  1000.0000 |         - | 7612.55 MB | 25.22x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [34] |     baseline |  14000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [34] | 2.51x slower | 115000.0000 |  1000.0000 |         - |  2078.6 MB |  7.93x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [34] | 1.61x slower |  41000.0000 |          - |         - |  739.22 MB |  2.82x more |
| ExcelFastXlsb                | sampl(...).xlsb [34] | 1.63x slower |  24000.0000 |          - |         - |  434.04 MB |  1.66x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] | 5.14x slower | 231000.0000 |  1000.0000 |         - | 4157.75 MB | 15.86x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [30] |     baseline |  14000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [30] | 2.46x slower | 115000.0000 |  1000.0000 |         - |  2078.6 MB |  7.93x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [30] | 1.56x slower |  41000.0000 |          - |         - |  739.22 MB |  2.82x more |
| ExcelFastXlsb                | sampl(...).xlsb [30] | 1.62x slower |  24000.0000 |          - |         - |  434.04 MB |  1.66x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] | 4.96x slower | 231000.0000 |  1000.0000 |         - | 4157.75 MB | 15.86x more |
```

## 2026-11-13 - V6 - Excel_FAST - Beta
- ✅ Add .Net11 TFM support
- Replaced the previous float round-trip memory comparison in CellValue.IsExactlyDouble with a fast memory-overlay check on the underlying 64-bit decimal mantissa (lo64 <= 999_999_999_999_999UL and hi == 0).
- Fixed XLSB pooled-record ownership so consumed records are returned promptly, rollback records remain valid until consumed or disposed, and parse/read failures do not strand rented buffers.

```
Job=net11.0  Runtime=.NET 11.0  IterationCount=3
LaunchCount=1  WarmupCount=3

| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  21000.0000 | 20000.0000 | 2000.0000 |  373.65 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.58x slower | 188000.0000 |  1000.0000 |         - | 3380.58 MB |  9.05x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.37x slower |  44000.0000 | 37000.0000 | 4000.0000 |  727.76 MB |  1.95x more |
| Excel_Prime              | 100mb.xlsx           | 1.25x slower |  39000.0000 | 36000.0000 | 4000.0000 |  649.58 MB |  1.74x more |
| ExcelFastXlsx            | 100mb.xlsx           | 1.09x slower |  40000.0000 | 37000.0000 | 5000.0000 |  649.61 MB |  1.74x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.26x slower |  51000.0000 | 39000.0000 | 4000.0000 |  865.92 MB |  2.32x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  12000.0000 |  1000.0000 |         - |  221.89 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.03x faster |  96000.0000 |  1000.0000 |         - | 1739.24 MB |  7.84x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.19x slower |  16000.0000 |  1000.0000 |         - |  294.92 MB |  1.33x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.05x slower |  12000.0000 |  1000.0000 |         - |  222.61 MB |  1.00x more |
| ExcelFastXlsx            | Blank(...).xlsx [30] | 1.16x faster |  12000.0000 |  1000.0000 |         - |  223.12 MB |  1.01x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.68x slower |  34000.0000 |  1000.0000 |         - |  620.05 MB |  2.79x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |   9000.0000 |          - |         - |  177.98 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.11x faster |  44000.0000 |          - |         - |  799.73 MB |  4.49x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.09x slower |  11000.0000 |  1000.0000 |         - |  213.65 MB |  1.20x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.03x faster |   9000.0000 |  1000.0000 |         - |  178.19 MB |  1.00x more |
| ExcelFastXlsx            | sampl(...).xlsx [34] | 1.27x faster |   9000.0000 |  1000.0000 |         - |  178.19 MB |  1.00x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.31x slower |  24000.0000 |  1000.0000 |         - |  440.81 MB |  2.48x more |
|                          |                      |              |             |            |           |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |   9000.0000 |          - |         - |  175.97 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.01x faster |  41000.0000 |          - |         - |  742.13 MB |  4.22x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.23x slower |  14000.0000 |  1000.0000 |         - |  252.71 MB |  1.44x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.15x slower |  12000.0000 |          - |         - |  217.05 MB |  1.23x more |
| ExcelFastXlsx            | sampl(...).xlsx [30] | 1.08x faster |  12000.0000 |          - |         - |  217.05 MB |  1.23x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.60x slower |  28000.0000 |  2000.0000 |         - |  513.74 MB |  2.92x more |
```
```
  ShortRun : .NET 8.0.34 (8.0.34, 8.0.3426.52217), X64 RyuJIT x86-64-v3
Job=ShortRun  IterationCount=3  LaunchCount=1  WarmupCount=3

| Method                       | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|----------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                    | 100mb.xlsb           |     baseline |  19000.0000 | 18000.0000 | 4000.0000 |  335.54 MB |             |
| AsyncExcel_PrimeXlsb         | 100mb.xlsb           | 2.36x slower | 167000.0000 | 33000.0000 | 4000.0000 | 2947.18 MB |  8.78x more |
| Excel_PrimeXlsb              | 100mb.xlsb           | 1.65x slower |  52000.0000 | 25000.0000 | 4000.0000 |  871.76 MB |  2.60x more |
| ExcelFastXlsb                | 100mb.xlsb           | 1.61x slower |  51000.0000 | 24000.0000 | 4000.0000 |  870.43 MB |  2.59x more |
| PrlAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           | 3.75x slower | 313000.0000 | 43000.0000 | 5000.0000 |  5541.4 MB | 16.51x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | Blank(...).xlsb [30] |     baseline |  16000.0000 |  1000.0000 |         - |  301.86 MB |             |
| AsyncExcel_PrimeXlsb         | Blank(...).xlsb [30] | 3.21x slower | 183000.0000 |  1000.0000 |         - | 3290.86 MB | 10.90x more |
| Excel_PrimeXlsb              | Blank(...).xlsb [30] | 2.01x slower |  45000.0000 |  1000.0000 |         - |  812.26 MB |  2.69x more |
| ExcelFastXlsb                | Blank(...).xlsb [30] | 1.89x slower |  45000.0000 |  1000.0000 |         - |  812.26 MB |  2.69x more |
| PrlAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] | 6.69x slower | 366000.0000 |  1000.0000 |         - | 6584.33 MB | 21.81x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [34] |     baseline |  14000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [34] | 2.59x slower | 102000.0000 |  1000.0000 |         - | 1842.06 MB |  7.02x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [34] | 1.69x slower |  28000.0000 |          - |         - |   502.7 MB |  1.92x more |
| ExcelFastXlsb                | sampl(...).xlsb [34] | 1.62x slower |  28000.0000 |          - |         - |   502.7 MB |  1.92x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] | 5.23x slower | 205000.0000 |  1000.0000 |         - | 3684.71 MB | 14.05x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [30] |     baseline |  14000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [30] | 2.51x slower | 102000.0000 |  1000.0000 |         - | 1842.06 MB |  7.02x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [30] | 1.70x slower |  28000.0000 |          - |         - |   502.7 MB |  1.92x more |
| ExcelFastXlsb                | sampl(...).xlsb [30] | 1.58x slower |  28000.0000 |          - |         - |   502.7 MB |  1.92x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] | 5.18x slower | 205000.0000 |  1000.0000 |         - | 3684.71 MB | 14.05x more |
```
.NET SDK 11.0.100
  [Host] : .NET 10.0.15 (10.0.15, 10.0.1526.53008), X64 RyuJIT x86-64-v3

Job=InProcess  Toolchain=InProcessEmitToolchain  MaxIterationCount=6
MinIterationCount=1  WarmupCount=2

| Method            | Mean     | Error   | Ratio | Allocated   | Alloc Ratio |
|------------------ |---------:|--------:|------:|------------:|------------:|
| Baseline          | 103.5 ms | 1.19 ms |  1.00 |   249.36 KB |        1.00 |
| SylvanXlsx        | 149.5 ms | 2.04 ms |  1.44 |   667.18 KB |        2.68 |
| PrimeFastXlsxType | 154.1 ms | 1.49 ms |  1.49 |  2758.31 KB |       11.06 |
| SylvanXlsx_BindT  | 157.6 ms | 2.54 ms |  1.52 | 11933.15 KB |       47.86 |
| PrimeFastXlsxObj  | 157.8 ms | 1.43 ms |  1.52 | 16971.75 KB |       68.06 |
| SylvanXlsxObj     | 163.4 ms | 1.52 ms |  1.58 | 14474.02 KB |       58.04 |
| PrimeXlsxType     | 183.1 ms | 3.29 ms |  1.77 |  2757.26 KB |       11.06 |
| PrimeXlsxObj      | 188.5 ms | 3.27 ms |  1.82 | 16970.51 KB |       68.06 |
| XlsxHelperXlsx    | 197.5 ms | 1.38 ms |  1.91 | 93306.14 KB |      374.18 |

