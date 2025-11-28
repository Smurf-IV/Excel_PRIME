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
-----

# 2025-10-18 pm
- All Cell Access
- Some performance Analysis and changes
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |--------------|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      | baseline     |  42000.0000 |  40000.0000 |   5000.0000 |  334.79 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.24x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB | 10.10x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 2.25x slower | 372000.0000 | 362000.0000 |   5000.0000 | 2940.45 MB |  8.78x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 2.23x slower | 363000.0000 | 353000.0000 |   5000.0000 | 2863.75 MB |  8.55x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] | baseline     | 392000.0000 | 376000.0000 | 375000.0000 | 2696.74 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.01x faster | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.90x slower | 423000.0000 | 422000.0000 |   5000.0000 | 3382.58 MB |  1.25x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.92x slower | 415000.0000 | 414000.0000 |   2000.0000 | 3312.35 MB |  1.23x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] | baseline     |  13000.0000 |           - |           - |  106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.05x slower | 100000.0000 |           - |           - |  799.74 MB |  7.49x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.98x slower | 201000.0000 | 200000.0000 |   3000.0000 | 1607.24 MB | 15.05x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 2.04x slower | 197000.0000 | 196000.0000 |   2000.0000 | 1572.79 MB | 14.73x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] | baseline     |  13000.0000 |           - |           - |  104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.09x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 2.08x slower | 194000.0000 | 193000.0000 |   3000.0000 | 1548.41 MB | 14.78x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 2.13x slower | 190000.0000 | 189000.0000 |   2000.0000 | 1514.05 MB | 14.45x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-10-19 pm
- Removed the "ToString()" call in the Benchmark, because the others did not have it!
- More profilling to see where things can be saved
- Now With `FastExcel` (Will not be using it again in these tests!)
```
| Method                          | FileName             | Ratio        | Gen0         | Gen1        | Gen2        | Allocated   | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |   42000.0000 |  40000.0000 |   5000.0000 |   334.79 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.35x slower |  424000.0000 |   5000.0000 |   2000.0000 |  3380.59 MB | 10.10x more |
| AccessEveryCellFastExcel        | Data/100mb.xlsx      | 8.26x slower | 3164000.0000 | 477000.0000 |   9000.0000 | 25311.91 MB | 75.61x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.76x slower |  457000.0000 |  34000.0000 |   5000.0000 |  3622.62 MB | 10.82x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.79x slower |  448000.0000 |  34000.0000 |   5000.0000 |  3545.91 MB | 10.59x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  393000.0000 | 377000.0000 | 376000.0000 |  2696.74 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.02x slower |  218000.0000 |   1000.0000 |           - |  1739.24 MB |  1.55x less |
| AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 8.94x slower | 6977000.0000 | 898000.0000 |  10000.0000 | 55575.12 MB | 20.61x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.51x slower |  572000.0000 |   2000.0000 |   1000.0000 |  4566.56 MB |  1.69x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.58x slower |  564000.0000 |   2000.0000 |   1000.0000 |   4496.4 MB |  1.67x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |   13000.0000 |           - |           - |   106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.02x slower |  100000.0000 |           - |           - |   799.73 MB |  7.49x more |
| AccessEveryCellFastExcel        | Data/(...).xlsx [39] | 6.69x slower | 1245000.0000 | 400000.0000 |   8000.0000 |  9874.29 MB | 92.48x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.51x slower |  268000.0000 |   1000.0000 |           - |  2141.68 MB | 20.06x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.60x slower |  264000.0000 |   1000.0000 |           - |  2107.34 MB | 19.74x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |   13000.0000 |           - |           - |   104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.10x slower |   93000.0000 |           - |           - |   742.13 MB |  7.08x more |
| AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 7.22x slower | 1137000.0000 | 292000.0000 |   8000.0000 |  9009.33 MB | 86.01x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.60x slower |  261000.0000 |   1000.0000 |           - |  2082.97 MB | 19.89x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.64x slower |  257000.0000 |   1000.0000 |           - |   2048.6 MB | 19.56x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
```
-----

# 2025-10-23
- After the fixes in the LazyStringLoader (Now appear to have more memory used??)
```
| Method                          | FileName             | Ratio        | Gen0         | Gen1        | Gen2        | Allocated   | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |   42000.0000 |  40000.0000 |   5000.0000 |   334.78 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.42x slower |  424000.0000 |   5000.0000 |   2000.0000 |  3380.59 MB | 10.10x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.85x slower |  564000.0000 |  56000.0000 |   6000.0000 |  4461.44 MB | 13.33x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.91x slower |  555000.0000 |  56000.0000 |   6000.0000 |  4384.73 MB | 13.10x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  391000.0000 | 375000.0000 | 374000.0000 |  2696.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.03x slower |  218000.0000 |   1000.0000 |           - |  1739.24 MB |  1.55x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.46x slower |  572000.0000 |   1000.0000 |           - |  4567.85 MB |  1.69x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.54x slower |  564000.0000 |   2000.0000 |   1000.0000 |  4497.72 MB |  1.67x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |   13000.0000 |           - |           - |   106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.03x slower |  100000.0000 |           - |           - |   799.73 MB |  7.49x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.43x slower |  268000.0000 |   1000.0000 |           - |  2141.85 MB | 20.06x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.48x slower |  264000.0000 |   1000.0000 |           - |  2107.51 MB | 19.74x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |   13000.0000 |           - |           - |   104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.10x slower |   93000.0000 |           - |           - |   742.13 MB |  7.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.53x slower |  261000.0000 |   1000.0000 |           - |  2083.12 MB | 19.89x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.56x slower |  257000.0000 |   1000.0000 |           - |  2048.75 MB | 19.56x more |
|-------------------------------- |--------------------- |-------------:|-------------:|------------:|------------:|------------:|------------:|
```
-----

# 2025-10-26
- Investigate the usage of a derived `XmlNameTable`
- Some tweaks to lifetimes
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  42000.0000 |  40000.0000 |   5000.0000 |  334.79 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.36x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB | 10.10x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.85x slower | 564000.0000 |  56000.0000 |   6000.0000 | 4461.44 MB | 13.33x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.93x slower | 555000.0000 |  56000.0000 |   6000.0000 | 4384.75 MB | 13.10x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 394000.0000 | 378000.0000 | 377000.0000 | 2696.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.04x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.47x slower | 572000.0000 |   1000.0000 |           - | 4567.98 MB |  1.69x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.49x slower | 564000.0000 |   2000.0000 |   1000.0000 | 4497.72 MB |  1.67x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  13000.0000 |           - |           - |  106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.05x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.42x slower | 268000.0000 |   1000.0000 |           - | 2141.85 MB | 20.06x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.55x slower | 264000.0000 |   1000.0000 |           - | 2107.63 MB | 19.74x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  13000.0000 |           - |           - |  104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.09x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.47x slower | 261000.0000 |   1000.0000 |           - | 2083.12 MB | 19.89x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.52x slower | 257000.0000 |   1000.0000 |           - | 2048.75 MB | 19.56x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-10-25
- make use of NameTable and `Object.ReferenceEquals`
- Allow many threads to add to the SharedStrings
- Add basic rawValue type detection
- Lowered the memory footprint
``` 
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  42000.0000 |  40000.0000 |   5000.0000 |   334.8 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.27x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB | 10.10x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 2.02x slower | 529000.0000 |  58000.0000 |   6000.0000 | 4188.77 MB | 12.51x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 2.05x slower | 520000.0000 |  56000.0000 |   6000.0000 | 4112.05 MB | 12.28x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 394000.0000 | 378000.0000 | 377000.0000 | 2696.74 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.03x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.72x slower | 509000.0000 |   2000.0000 |   1000.0000 | 4065.55 MB |  1.51x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.79x slower | 501000.0000 |   2000.0000 |   1000.0000 | 3995.39 MB |  1.48x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  13000.0000 |           - |           - |  106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.04x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.75x slower | 237000.0000 |   1000.0000 |           - | 1891.65 MB | 17.72x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.90x slower | 233000.0000 |   1000.0000 |           - | 1857.28 MB | 17.39x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  13000.0000 |           - |           - |  104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.08x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.73x slower | 244000.0000 |   1000.0000 |           - | 1950.46 MB | 18.62x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.80x slower | 240000.0000 |   1000.0000 |           - | 1915.95 MB | 18.29x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-10-26
- Make use of `RestrictedNameTable`s
- Use additional offset to reduce locking intensity
- Use Atoms in the Sheet reader
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  42000.0000 |  40000.0000 |   5000.0000 |  334.79 MB |         |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.38x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB | 10.10x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.91x slower | 518000.0000 |  53000.0000 |   6000.0000 | 4093.55 MB | 12.23x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.96x slower | 508000.0000 |  54000.0000 |   6000.0000 | 4016.82 MB | 12.00x more |
|                                 |                      |              |             |             |             |            |             |
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 392000.0000 | 376000.0000 | 375000.0000 | 2696.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.05x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.58x slower | 509000.0000 |   2000.0000 |   1000.0000 | 4065.25 MB |  1.51x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.68x slower | 501000.0000 |   2000.0000 |   1000.0000 | 3995.09 MB |  1.48x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  13000.0000 |           - |           - |  106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.07x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.62x slower | 237000.0000 |   1000.0000 |           - | 1891.62 MB | 17.72x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.73x slower | 233000.0000 |   1000.0000 |           - | 1857.23 MB | 17.39x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  13000.0000 |           - |           - |  104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.10x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.64x slower | 244000.0000 |   1000.0000 |           - | 1950.32 MB | 18.62x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.71x slower | 240000.0000 |   1000.0000 |           - | 1915.94 MB | 18.29x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-11-01
- Make all the benchmarks perform a `ToString` to ensure data is actually retrieved
- Pass `Atomized` strings around
- Pass `InstanceContext` around
- Add `None` as the default convert option
- Remove File buffers to allow OS to perform caching
- Correct Sheet Id's
- Extend options for Value extraction
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  43000.0000 |  41000.0000 |   5000.0000 |  338.74 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.03x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.82x slower | 525000.0000 |  55000.0000 |   6000.0000 | 4153.29 MB | 12.26x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.88x slower | 516000.0000 |  53000.0000 |   6000.0000 | 4076.56 MB | 12.03x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 408000.0000 | 370000.0000 | 369000.0000 | 2875.5 MB  |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.03x faster | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.65x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.48x slower | 556000.0000 |   2000.0000 |   1000.0000 | 4436.84 MB |  1.54x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.53x slower | 548000.0000 |   2000.0000 |   1000.0000 | 4366.69 MB |  1.52x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  33000.0000 |           - |           - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.10x faster | 100000.0000 |           - |           - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.33x slower | 284000.0000 |   1000.0000 |           - | 2271.62 MB |  8.55x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.45x slower | 280000.0000 |   1000.0000 |           - | 2237.21 MB |  8.42x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] | baseline     |  36000.0000 |           - |           - |  294.17 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.08x faster |  93000.0000 |           - |           - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.49x slower | 252000.0000 |   1000.0000 |           - | 2015.16 MB |  6.85x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.51x slower | 248000.0000 |   1000.0000 |           - | 1980.78 MB |  6.73x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-11-04
- Move Rented buffer into Row
- Use `char[]` for ColumnName storage
- Tinker with `DefinedRange` class
- Add styles of extraction to benchmark
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      | baseline     |  43000.0000 |  41000.0000 |   5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.19x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.86x slower | 508000.0000 |  55000.0000 |   6000.0000 | 4019.48 MB | 11.87x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.88x slower | 499000.0000 |  56000.0000 |   6000.0000 | 3942.68 MB | 11.64x more |
| SimpleCellAsyncExcel_Prime      | Data/100mb.xlsx      | 1.84x slower | 508000.0000 |  55000.0000 |   6000.0000 | 4019.48 MB | 11.87x more |
| NumberCellAsyncExcel_Prime      | Data/100mb.xlsx      | 1.79x slower | 508000.0000 |  55000.0000 |   6000.0000 | 4019.46 MB | 11.87x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] | baseline     | 426000.0000 | 388000.0000 | 387000.0000 | 2875.49 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.04x faster | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.65x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.43x slower | 526000.0000 |   2000.0000 |   1000.0000 | 4202.51 MB |  1.46x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.48x slower | 518000.0000 |   2000.0000 |   1000.0000 | 4131.86 MB |  1.44x more |
| SimpleCellAsyncExcel_Prime      | Data/(...).xlsx [35] | 1.47x slower | 526000.0000 |   2000.0000 |   1000.0000 | 4202.44 MB |  1.46x more |
| NumberCellAsyncExcel_Prime      | Data/(...).xlsx [35] | 1.45x slower | 503000.0000 |   2000.0000 |   1000.0000 | 4016.72 MB |  1.40x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [39] | baseline     |  33000.0000 |           - |           - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.09x faster | 100000.0000 |           - |           - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.35x slower | 271000.0000 |   1000.0000 |           - | 2166.13 MB |  8.15x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.42x slower | 267000.0000 |   1000.0000 |           - | 2131.77 MB |  8.02x more |
| SimpleCellAsyncExcel_Prime      | Data/(...).xlsx [39] | 1.29x slower | 271000.0000 |   1000.0000 |           - | 2166.13 MB |  8.15x more |
| NumberCellAsyncExcel_Prime      | Data/(...).xlsx [39] | 1.40x slower | 247000.0000 |   1000.0000 |           - | 1971.27 MB |  7.42x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/(...).xlsx [35] | baseline     |  36000.0000 |           - |           - |  294.17 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.07x faster |  93000.0000 |           - |           - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.46x slower | 239000.0000 |   1000.0000 |           - | 1909.71 MB |  6.49x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.50x slower | 235000.0000 |   1000.0000 |           - | 1875.26 MB |  6.37x more |
| SimpleCellAsyncExcel_Prime      | Data/(...).xlsx [35] | 1.45x slower | 239000.0000 |   1000.0000 |           - | 1909.71 MB |  6.49x more |
| NumberCellAsyncExcel_Prime      | Data/(...).xlsx [35] | 1.46x slower | 239000.0000 |   1000.0000 |           - | 1909.65 MB |  6.49x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-11-07
- Add option to use the Stream from `ZipEntry`
- Explain the Excel cell types
- Make the Benchmark "File Names" clearer
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 |  41000.0000 |   5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.26x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.60x slower | 507000.0000 |  55000.0000 |   6000.0000 | 4012.69 MB | 11.85x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.57x slower | 498000.0000 |  55000.0000 |   6000.0000 | 3936.14 MB | 11.62x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.73x slower | 507000.0000 |  55000.0000 |   6000.0000 | 4012.67 MB | 11.85x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline | 419000.0000 | 381000.0000 | 380000.0000 | 2875.49 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.03x faster | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.65x less |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.21x slower | 524000.0000 |   1000.0000 |           - |    4187 MB |  1.46x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.15x slower | 516000.0000 |   1000.0000 |           - | 4117.17 MB |  1.43x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.25x slower | 501000.0000 |   1000.0000 |           - | 4001.28 MB |  1.39x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |           - |           - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.11x faster | 100000.0000 |           - |           - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.13x slower | 270000.0000 |   1000.0000 |           - | 2160.73 MB |  8.13x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.10x slower | 266000.0000 |   1000.0000 |           - |  2126.5 MB |  8.00x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.23x slower | 246000.0000 |   1000.0000 |           - | 1965.82 MB |  7.40x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |           - |           - |  294.17 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.08x faster |  93000.0000 |           - |           - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.26x slower | 238000.0000 |   1000.0000 |           - | 1905.92 MB |  6.48x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.23x slower | 234000.0000 |   1000.0000 |           - | 1871.72 MB |  6.36x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.29x slower | 238000.0000 |   1000.0000 |           - | 1905.92 MB |  6.48x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-11-08
- Investigation into the smallest function ;-)
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 |  41000.0000 |   5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.05x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.58x slower | 507000.0000 |  55000.0000 |   6000.0000 | 4012.67 MB | 11.85x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.53x slower | 498000.0000 |  55000.0000 |   6000.0000 | 3936.15 MB | 11.62x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.57x slower | 504000.0000 |  52000.0000 |   3000.0000 | 4012.66 MB | 11.85x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline | 399000.0000 | 361000.0000 | 360000.0000 | 2875.46 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.05x faster | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.65x less |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.22x slower | 524000.0000 |   1000.0000 |           - |    4187 MB |  1.46x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.16x slower | 516000.0000 |   1000.0000 |           - | 4117.18 MB |  1.43x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.23x slower | 501000.0000 |   1000.0000 |           - | 4001.28 MB |  1.39x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |           - |           - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.12x faster | 100000.0000 |           - |           - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.12x slower | 270000.0000 |   1000.0000 |           - | 2160.73 MB |  8.13x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.10x slower | 266000.0000 |   1000.0000 |           - |  2126.5 MB |  8.00x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.21x slower | 246000.0000 |   1000.0000 |           - | 1965.84 MB |  7.40x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |           - |           - |  294.17 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.08x faster |  93000.0000 |           - |           - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.27x slower | 238000.0000 |   1000.0000 |           - | 1905.91 MB |  6.48x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.23x slower | 234000.0000 |   1000.0000 |           - | 1871.72 MB |  6.36x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.25x slower | 238000.0000 |   1000.0000 |           - | 1905.91 MB |  6.48x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-11-12
- Optimise for `CellConversion.None`

```
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 |  41000.0000 |   5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.08x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.60x slower | 521000.0000 |  53000.0000 |   6000.0000 | 4123.51 MB | 12.17x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.54x slower | 512000.0000 |  54000.0000 |   6000.0000 |  4046.9 MB | 11.95x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.58x slower | 521000.0000 |  53000.0000 |   6000.0000 |  4123.5 MB | 12.17x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline | 392000.0000 | 354000.0000 | 353000.0000 | 2875.49 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.02x faster | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.65x less |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.20x slower | 501000.0000 |   1000.0000 |           - | 4002.64 MB |  1.39x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.15x slower | 493000.0000 |   1000.0000 |           - | 3932.82 MB |  1.37x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.25x slower | 514000.0000 |   1000.0000 |           - | 4102.53 MB |  1.43x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |           - |           - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.09x faster | 100000.0000 |           - |           - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.09x slower | 237000.0000 |           - |           - | 1894.17 MB |  7.13x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.07x slower | 233000.0000 |   1000.0000 |           - | 1859.94 MB |  7.00x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.24x slower | 252000.0000 |   1000.0000 |           - | 2015.43 MB |  7.59x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |           - |           - |  294.17 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.09x faster |  93000.0000 |           - |           - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.16x slower | 240000.0000 |   1000.0000 |           - | 1917.02 MB |  6.52x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.10x slower | 236000.0000 |   1000.0000 |           - | 1882.83 MB |  6.40x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.28x slower | 245000.0000 |   1000.0000 |           - | 1955.51 MB |  6.65x more |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
```
-----

# 2025-11-16 - Beta V2
- Add `IEnumerable`s _All_ the way down ⤵️
    - i.e. remove the need for Asynchronous awaits
    - ⛓️‍💥 Breaking Change 🔩
        - The Async classes now have `Async` appended to be distinct from the non async versions
        - But, `Async` inherit from the non, so they are interchangable

```
| Method                          | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 | 5000.0000 |  338.73 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.05x slower | 424000.0000 |  5000.0000 | 2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.57x slower | 523000.0000 | 54000.0000 | 6000.0000 | 4137.73 MB | 12.22x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.50x slower | 503000.0000 | 57000.0000 | 6000.0000 | 3972.08 MB | 11.73x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.57x slower | 520000.0000 | 52000.0000 | 3000.0000 | 4139.47 MB | 12.22x more |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 | 1000.0000 |  318.93 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.01x faster | 218000.0000 |  1000.0000 |         - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.26x slower | 506000.0000 |  1000.0000 |         - | 4037.46 MB | 12.66x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.20x slower | 488000.0000 |  1000.0000 |         - | 3897.37 MB | 12.22x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.36x slower | 521000.0000 |  1000.0000 |         - | 4160.83 MB | 13.05x more |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |         - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.11x faster | 100000.0000 |          - |         - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.07x slower | 237000.0000 |  1000.0000 |         - | 1893.49 MB |  7.13x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.06x slower | 228000.0000 |  1000.0000 |         - |    1825 MB |  6.87x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.29x slower | 252000.0000 |  1000.0000 |         - | 2014.79 MB |  7.58x more |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |         - |  294.16 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.07x faster |  93000.0000 |          - |         - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.19x slower | 240000.0000 |  1000.0000 |         - | 1916.81 MB |  6.52x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.08x slower | 231000.0000 |  1000.0000 |         - | 1848.26 MB |  6.28x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.40x slower | 257000.0000 |  1000.0000 |         - | 2053.52 MB |  6.98x more |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
```
-----

# 2025-11-25 - Beta V2
- Make `DefinedName`'s work with `localSheetId`definitions
- Start to Add Benchmarks for range extraction

```
| Method      | ranger               | Mean     | Error     | StdDev   | Gen0         | Gen1        | Gen2       | Allocated |
|------------ |--------------------- |---------:|----------:|---------:|-------------:|------------:|-----------:|----------:|
| Access100mb | Excel(...)edXML [39] | 39.508 s | 10.2098 s | 0.5596 s | 1248000.0000 | 382000.0000 |  6000.0000 |  10.21 GB |
| Access100mb | Excel(...)PPlus [36] | 13.795 s |  2.9397 s | 0.1611 s |  593000.0000 | 163000.0000 | 18000.0000 |   7.33 GB |
| Access100mb | Excel(...)Prime [40] |  8.079 s |  0.9203 s | 0.0504 s |  244000.0000 | 127000.0000 |  7000.0000 |   1.86 GB |
```
-----

# 2025-11-27 - Beta V2
- Change benchmarks to use `ToString` to be fair on `ClosedXML`
- Attempt to use `FastExcel` - Failed due to Bug
- Use _commercial_ `FreeSpire`

```
| Method      | ranger               | Mean     | Error    | StdDev   | Gen0         | Gen1        | Gen2       | Allocated |
|------------ |--------------------- |---------:|---------:|---------:|-------------:|------------:|-----------:|----------:|
| Access100mb | Excel(...)edXML [39] | 41.751 s | 3.9483 s | 0.2164 s | 1248000.0000 | 382000.0000 |  6000.0000 |  10.21 GB |
| Access100mb | Excel(...)PPlus [36] | 14.232 s | 2.9999 s | 0.1644 s |  593000.0000 | 163000.0000 | 18000.0000 |   7.33 GB |
| Access100mb | Excel(...)Prime [40] |  8.145 s | 0.9088 s | 0.0498 s |  244000.0000 | 127000.0000 |  7000.0000 |   1.86 GB |
| Access100mb | Excel(...)Spire [39] | 24.023 s | 9.0692 s | 0.4971 s | 1181000.0000 | 445000.0000 |  6000.0000 |   9.95 GB |
```
-----

# 2025-11-18 - Beta V2
- Implement usage of _commercial_ `Aspose.Cells`
- Switch to `EPPlus-LPGL` (It's faster than v8.x ;-))

```
| Method      | ranger               | Mean     | Error    | StdDev   | Gen0         | Gen1        | Gen2       | Allocated |
|------------ |--------------------- |---------:|---------:|---------:|-------------:|------------:|-----------:|----------:|
| Access100mb | Excel(...)Cells [41] |  7.743 s | 0.1866 s | 0.0102 s |  175000.0000 |  98000.0000 |  7000.0000 |   1.49 GB |
| Access100mb | Excel(...)edXML [39] | 40.052 s | 2.4428 s | 0.1339 s | 1248000.0000 | 382000.0000 |  6000.0000 |  10.21 GB |
| Access100mb | Excel(...)PPlus [36] | 11.381 s | 1.1321 s | 0.0621 s |  550000.0000 | 189000.0000 | 13000.0000 |   6.08 GB |
| Access100mb | Excel(...)Prime [40] |  7.786 s | 0.3911 s | 0.0214 s |  244000.0000 | 127000.0000 |  7000.0000 |   1.86 GB |
| Access100mb | Excel(...)Spire [39] | 23.528 s | 1.8896 s | 0.1036 s | 1181000.0000 | 449000.0000 |  6000.0000 |   9.95 GB |
```
-----