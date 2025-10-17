<!--TOC-->
- [Intro](#intro)
- [2025-10-08](#2025-10-08)
- [2025-10-13](#2025-10-13)
- [2025-10-14](#2025-10-14)
<!--/TOC-->

# Intro
All done with the following
```
BenchmarkDotNet v0.15.4, Windows 11 (10.0.26100.6584/24H2/2024Update/HudsonValley)
Intel Core i9-9900K CPU 3.60GHz (Coffee Lake), 1 CPU, 16 logical and 8 physical cores
```
And then slightly different versions of the following dependent on date:
```
.NET SDK 10.0.100-rc.1.25451.107
  [Host] : .NET 8.0.20 (8.0.20, 8.0.2025.41914), X64 RyuJIT x86-64-v3
```
<hr />

# 2025-10-08
- Can the large files be loaded

| Method                     | Max          |
|--------------------------- |-------------:|
| **AccessEveryCellSylvan**      |  **4,835.00 ms** |
| AccessEveryCellXlsxHelper  | 20,724.69 ms |
| AccessEveryCellExcel_Prime | 33,230.03 ms |
|                            |              |
| **AccessEveryCellSylvan**      |  **7,339.47 ms** |
| AccessEveryCellXlsxHelper  |  8,325.41 ms |
| AccessEveryCellExcel_Prime | 57,440.71 ms |
|                            |              |
| **AccessEveryCellSylvan**      |  **3,184.49 ms** |
| AccessEveryCellXlsxHelper  |  3,945.23 ms |
| AccessEveryCellExcel_Prime | 25,463.68 ms |
|                            |              |
| **AccessEveryCellSylvan**      |     **77.64 ms** |
| AccessEveryCellXlsxHelper  |  3,551.77 ms |
| AccessEveryCellExcel_Prime | 21,429.76 ms |

<hr />

# 2025-10-13
- This set of tests "Only" accessed the first cel of the returned rows
- Therefore did not really excercise the retrieval of the `SharedStrings`

| Method                          | FileName             | Mean         | Error        | StdDev       | Ratio            | RatioSD | Gen0         | Gen1         | Gen2       | Allocated   | Alloc Ratio  |
|-------------------------------- |--------------------- |-------------:|-------------:|-------------:|-----------------:|--------:|-------------:|-------------:|-----------:|------------:|-------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |  4,429.30 ms |    981.52 ms |    53.800 ms |         baseline |         |   49000.0000 |   47000.0000 |  3000.0000 |    403.8 MB |              |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 18,815.59 ms |  2,917.95 ms |   159.943 ms |     4.25x slower |   0.05x |  424000.0000 |    5000.0000 |  2000.0000 |  3380.59 MB |   8.37x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 31,856.62 ms | 20,032.72 ms | 1,098.061 ms |     7.19x slower |   0.23x | 1241000.0000 |  870000.0000 |  1000.0000 |  9935.96 MB |  24.61x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 32,497.09 ms |  6,566.78 ms |   359.948 ms |     7.34x slower |   0.10x | 1231000.0000 |  861000.0000 | 10000.0000 |  9782.58 MB |  24.23x more |
|                                 |                      |              |              |              |                  |         |              |              |            |             |              |
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |  7,454.71 ms |  1,003.66 ms |    55.014 ms |         baseline |         |   65000.0000 |   46000.0000 | 41000.0000 |  2736.67 MB |              |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] |  7,585.11 ms |    241.34 ms |    13.229 ms |     1.02x slower |   0.01x |  218000.0000 |    1000.0000 |          - |  1739.24 MB |   1.57x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 56,598.18 ms | 34,983.95 ms | 1,917.588 ms |     7.59x slower |   0.23x | 2276000.0000 | 1521000.0000 | 10000.0000 | 18080.92 MB |   6.61x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 58,418.49 ms |  6,662.92 ms |   365.217 ms |     7.84x slower |   0.07x | 2249000.0000 | 1492000.0000 |  1000.0000 | 17940.46 MB |   6.56x more |
|                                 |                      |              |              |              |                  |         |              |              |            |             |              |
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |  3,246.19 ms |     70.94 ms |     3.888 ms |         baseline |         |   14000.0000 |    1000.0000 |          - |   115.51 MB |              |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] |  3,415.80 ms |    119.77 ms |     6.565 ms |     1.05x slower |   0.00x |  100000.0000 |            - |          - |   799.73 MB |   6.92x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 24,282.55 ms |  9,097.37 ms |   498.658 ms |     7.48x slower |   0.13x | 1028000.0000 |  693000.0000 |  8000.0000 |  8141.35 MB |  70.48x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 24,936.06 ms |  5,544.87 ms |   303.933 ms |     7.68x slower |   0.08x | 1020000.0000 |  685000.0000 |  8000.0000 |  8072.55 MB |  69.88x more |
|                                 |                      |              |              |              |                  |         |              |              |            |             |              |
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     18.32 ms |     11.10 ms |     0.608 ms |         baseline |         |    2468.7500 |    2375.0000 |   218.7500 |    18.07 MB |              |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] |  3,051.97 ms |    257.79 ms |    14.131 ms |   166.70x slower |   4.79x |   93000.0000 |            - |          - |   742.14 MB |  41.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 21,005.90 ms |  2,875.97 ms |   157.641 ms | 1,147.33x slower |  33.46x |  821000.0000 |  590000.0000 |  8000.0000 |  6485.07 MB | 358.93x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 20,832.36 ms |  1,993.69 ms |   109.281 ms | 1,137.85x slower |  32.76x |  812000.0000 |  581000.0000 |  8000.0000 |  6416.41 MB | 355.13x more |

<hr />

# 2025-10-14
- This set of tests accessed the every cell of the returned rows
- Therefore excercises the retrieval of the all `SharedStrings`


| Method                          | FileName             | Mean         | Error         | StdDev       | Ratio            | RatioSD | Gen0         | Gen1         | Gen2       | Allocated   | Alloc Ratio  |
|-------------------------------- |--------------------- |-------------:|--------------:|-------------:|-----------------:|--------:|-------------:|-------------:|-----------:|------------:|-------------:|
| **AccessEveryCellSylvan**           | **Data/100mb.xlsx**      |  **4,552.13 ms** |    **860.207 ms** |    **47.151 ms** |         **baseline** |        **** |   **49000.0000** |   **47000.0000** |  **3000.0000** |    **403.8 MB** |             **** |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 19,588.22 ms |  2,998.794 ms |   164.374 ms |     4.30x slower |   0.05x |  424000.0000 |    5000.0000 |  2000.0000 |  3380.59 MB |   8.37x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 31,440.00 ms | 13,731.937 ms |   752.694 ms |     6.91x slower |   0.16x | 1236000.0000 |  865000.0000 |  1000.0000 |  9861.72 MB |  24.42x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 29,864.75 ms | 16,507.782 ms |   904.847 ms |     6.56x slower |   0.18x | 1225000.0000 |  854000.0000 | 10000.0000 |  9708.04 MB |  24.04x more |
|                                 |                      |              |               |              |                  |         |              |              |            |             |              |
| **AccessEveryCellSylvan**           | **Data/(...).xlsx [35]** |  **7,607.97 ms** |  **2,659.789 ms** |   **145.792 ms** |         **baseline** |        **** |   **65000.0000** |   **46000.0000** | **41000.0000** |  **2736.63 MB** |             **** |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] |  7,558.58 ms |    656.938 ms |    36.009 ms |     1.01x faster |   0.02x |  218000.0000 |    1000.0000 |          - |  1739.24 MB |   1.57x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 57,067.48 ms |  1,083.106 ms |    59.369 ms |     7.50x slower |   0.12x | 2267000.0000 | 1508000.0000 |  1000.0000 | 18080.79 MB |   6.61x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 58,214.50 ms | 18,809.976 ms | 1,031.038 ms |     7.65x slower |   0.17x | 2258000.0000 | 1503000.0000 | 10000.0000 | 17940.49 MB |   6.56x more |
|                                 |                      |              |               |              |                  |         |              |              |            |             |              |
| **AccessEveryCellSylvan**           | **Data/(...).xlsx [39]** |  **3,284.64 ms** |    **102.451 ms** |     **5.616 ms** |         **baseline** |        **** |   **14000.0000** |    **1000.0000** |          **-** |   **115.46 MB** |             **** |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] |  3,392.15 ms |    189.324 ms |    10.378 ms |     1.03x slower |   0.00x |  100000.0000 |            - |          - |   799.73 MB |   6.93x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 24,851.86 ms |  3,904.259 ms |   214.006 ms |     7.57x slower |   0.06x | 1028000.0000 |  692000.0000 |  8000.0000 |  8141.21 MB |  70.51x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 23,939.73 ms |  2,082.430 ms |   114.145 ms |     7.29x slower |   0.03x | 1020000.0000 |  684000.0000 |  8000.0000 |  8072.74 MB |  69.92x more |
|                                 |                      |              |               |              |                  |         |              |              |            |             |              |
| **AccessEveryCellSylvan**           | **Data/(...).xlsx [35]** |     **19.17 ms** |      **4.471 ms** |     **0.245 ms** |         **baseline** |        **** |    **2468.7500** |    **2375.0000** |   **218.7500** |    **18.07 MB** |             **** |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] |  2,984.95 ms |    474.120 ms |    25.988 ms |   155.76x slower |   2.08x |   93000.0000 |            - |          - |   742.13 MB |  41.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 21,909.18 ms | 15,112.848 ms |   828.386 ms | 1,143.25x slower |  39.51x |  821000.0000 |  590000.0000 |  8000.0000 |  6485.06 MB | 358.93x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 21,442.36 ms |  1,281.360 ms |    70.236 ms | 1,118.89x slower |  12.75x |  812000.0000 |  581000.0000 |  8000.0000 |  6416.39 MB | 355.13x more |

<hr />

# 2025-10-17
- Shows that the strings are now lazy loaded
- Still not as good as `XlsxHelper` for the just file loading
- `FastExcel` does not like these files - or I'm probably doing something wrong

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