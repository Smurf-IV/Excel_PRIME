# Excel_PRIME
**Excel**_**P**erformant **R**eader via **I**nterfaces for **M**emory **E**fficiency.

# What does that mean?
_Yet another Excel reader ?_, but starting with .Net 8 as the performant Runtime.

Lets take each of the above elements and explain:

## Excel
- Open _Large_ 2007 (Onwards) XLSX file formats (Binary later maybe)

## Performant
- Try to be as fast as possible, i.e.
    - Forward only Lazy loading
    - No Attempting to decipher / convert the cell(s) types (Its all text in lx)
    - No attempting to create /use datatables with headers etc.
    - Use `IEnumerable`s with initial offset starts (Row / Column)
    - Allow `CancellationToken`s to be used to allow page transitioning cancellation (More on this later)

## Reader
Read only, therefore no calculation / formula calls

## Interfaces 
- Will use the DotNet core functionality by default
- But, if your target deployment allows for the use of native performant binaries, then via the use of interfaces these will be pluggable
    - i.e. Using `Zlib.Net` for getting the data streams out of the compressed Excel file faster.
    - A faster / slimmer implementation for xml stream reading (i.e. TurboXml)

## Memory
- The reason for this is to handle very large XSLX files (i.e. > 500K rows with > 180 columns per sheet, with multiple sheets of this size)
- For `ETL` validation scenarios, i.e. make sure that the user modified data that has been transferred has interaction rules applied, before moving onto the `T` and `L` stages
- Try not to hit / store in the LOH
- No internal caching of previously loaded sheets / rows.

## Efficiency
- As hinted by the above statements, this is to be targetted at memory restricted environments (i.e. ASP Net VM's)
- Use the OS's `Temp File` caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
- Only unzip the sheet(s) when they are asked for

## Etc.
### `CancellationToken`s
- This is to allow the Large files to be _Aborted_
- Make "Most" of the API's Asynchronous `Task`s
### IDisposable
- Got to tidy up those `Temp File`s, and release the `File Stream`

<hr />

# It will **_not_** be:
## Thread safe
- Initially it will **Not** be _same sheet_ thread safe, because the xml reader will be locked to the sheet in use.
## Cell object type
- Cell converted when read (i.e. you will know the type that you want, and you can convert it.)
- This could later become an option if the `XmlConvert` classes are efficient (Or via the interface specs)
## Poco
- A POCO / Type populator (Extensions can be written for that later)
## Writer / Modifier
- Totally beyond the scope of this project remit

<hr />

| Badge   | Area   |
|--------------------------- |-------------|
| [![.NET](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml/badge.svg?branch=main)](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml) | Release build and tests |

<hr />

# Targets
## Phase 0
- [x] Setup this github
- [x] Create the main project
- [x] Add Unit Test project
- [x] Add simple Test Data

## Phase Alpha
- [x] Use Net Core Interface(s)
    - [x] Use `ZipArchive`
    - [x] Use `XDocument`
- [x] Implement Open / Dispose (Async)
    - [x] Sheet Names
    - [x] Shared Strings
- [x] Implement Sheet loading (unzip and be ready for use)
    - [x] Use `XDocument`
- [x] Implement Row extraction 
    - [x] Skip
    - [x] Delayed read - until a cell is actually needed
    - [x] Deal with Null / Empty cells (Utilise sparse array)
    - [x] Keep last used offset (i.e. no need to reload sheet if the next range API `startRow` call is later)

## Phase Beta - Benchmarks
- [x] Benchmarks
    - [x] Add Other "Excel readers" to the Benchmark project(s)
- [x] More UnitTests

Seems like I have some work to do:
```
BenchmarkDotNet v0.15.4, Windows 11 (10.0.26100.6584/24H2/2024Update/HudsonValley)
Intel Core i9-9900K CPU 3.60GHz (Coffee Lake), 1 CPU, 16 logical and 8 physical cores
.NET SDK 10.0.100-rc.1.25451.107
  [Host] : .NET 8.0.20 (8.0.20, 8.0.2025.41914), X64 RyuJIT x86-64-v3
  Dry    : .NET 8.0.20 (8.0.20, 8.0.2025.41914), X64 RyuJIT x86-64-v3
```
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

## Phase 1 - MVP
- [ ] Add Non `IAsyncEnumerable`s and benchmark
    - [x] Changes done 2025-10-13

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

    - [x] Changes done 2025-10-14

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



- [ ] - [ ] Implement `XmlReader.Create` for
    - [x] Loading sharedStrings
    - [ ] Sheet loading
- [x] Better `Storage` of the SharedStrings
    - [x] Use of LazyLoading Class
- [ ] More Benchmarks
- [ ] Read `definedName`s (Ranges)
    - [ ] Store from global
- [ ] Implement Sheet loading of
    - [ ] Q: Are there _Shared strings_ per sheet as well ?
    - [ ] Store `definedName` from Local sheets (When opened)
- [ ] Implement Row extraction 
    - [ ] Allow ColumnHeader addressing (i.e. `ABF`)
- [ ] Implement RangeExtraction
    - [ ] Global rangeNames
    - [ ] Per Sheet rangeNames
    - [ ] User defined

## Phase 2 - Multi project deployments (Nuget)
- [ ] More Benchmarks
    - [ ] Add even more "Excel readers" to the Benchmark project(s)
- [ ] Cell Information
    - [ ] Value type indicator
    - [ ] Formatter applied
    - [ ] Anything else useful
- [ ] Investigate a different way of storing the _Shared strings_ to the Filesystem, when they are in the MB's
- [ ] Implement Interface for other Libs (Xml / Zip)

## Phase 3
- [ ] XLS**B**


## Phase 4 - Extension(s)
- [ ] Add option class to allow _Basic_ Cell value type identification
    - [ ] Extract into those types
    - [ ] Deal with `DateOnly` / `TimeOnly` fields
    - [ ] Use of user defined column schema (Excel Number Format nuget?)
- [ ] Investigate possibility of using "Pipelining" to get data for Next row / cell population after yield?
    - [ ] Locking
    - [ ] How to deal with rows that are completely blank
    - [ ] `fibres` ?

- [ ] More ideas to be added later ;-)
