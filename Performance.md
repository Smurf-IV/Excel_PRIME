<!-- TOC-->
- [Intro](#intro)
- [2025-10-18 pm](#2025-10-18-pm)
- [2025-10-19 pm](#2025-10-19-pm)
- [2025-10-23](#2025-10-23)
- [2025-10-26](#2025-10-26)
- [2025-10-25](#2025-10-25)
- [2025-10-26](#2025-10-26-1)
- [2025-11-01](#2025-11-01)
- [2025-11-04](#2025-11-04)
- [2025-11-07](#2025-11-07)
- [2025-11-08](#2025-11-08)
- [2025-11-12](#2025-11-12)
- [2025-11-16 - Beta V2](#2025-11-16---beta-v2)
- [2025-11-25 - Beta V2](#2025-11-25---beta-v2)
- [2025-11-27 - Beta V2](#2025-11-27---beta-v2)
- [2025-11-18 - Beta V2](#2025-11-18---beta-v2)
- [2025-12-01 - Beta V2](#2025-12-01---beta-v2)
- [2025-12-02 - Beta V2](#2025-12-02---beta-v2)
- [2025-12-04 - Beta V2](#2025-12-04---beta-v2)
- [2025-12-07 - Beta V2](#2025-12-07---beta-v2)
- [2025-12-10](#2025-12-10)
- [2025-12-20](#2025-12-20)
- [2025-12-21](#2025-12-21)
- [2026-01-02](#2026-01-02)
- [2026-01-04](#2026-01-04)
- [2026-01-11](#2026-01-11)
- [2026-01-16 - V3](#2026-01-16---v3)
- [2026-01-18 - V4 Alpha](#2026-01-18---v4-alpha)
- [2026-01-20 - V4 Alpha](#2026-01-20---v4-alpha)
- [2026-01-31 - V4 Alpha](#2026-01-31---v4-alpha)
- [2026-04-05 - V4 Beta](#2026-04-05---v4-beta)
- [2026-04-10 - V4 - Beta](#2026-04-10---v4---beta)
- [2026-04-22 - V4 - RC2](#2026-04-22---v4---rc2)
- [2026-04-27 - V4 - RC3](#2026-04-27---v4---rc3)
- [2026-05-05 - V4](#2026-05-05---v4)
- [2026-05-12 - V4 - Bug fixes](#2026-05-12---v4---bug-fixes)
- [2026-05-25 - V5 - Beta](#2026-05-25---v5---beta)
<!-- TOC -->

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

# 2025-12-01 - Beta V2
- Replaced the dictionary with a fixed-size Cell?[] 
- Defensive bounds checks when assigning parsed cells to avoid out-of-range writes.
- `Row` disposal when going out of `yield` scopes
- Add `ThreadStringBuilderPool` for memory efficiency
- Add `AccessPivotTable` and explain why the other libraries do not work

```
| Method                          | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.10x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.55x slower | 413000.0000 | 50000.0000 |  3000.0000 | 3286.11 MB |  9.70x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.27x slower | 318000.0000 | 54000.0000 |  6000.0000 | 2504.76 MB |  7.39x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.59x slower | 417000.0000 | 53000.0000 |  6000.0000 | 3287.89 MB |  9.71x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.02x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.22x slower | 458000.0000 | 42000.0000 | 41000.0000 |  3455.5 MB | 10.84x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.08x faster | 307000.0000 | 42000.0000 | 41000.0000 | 2250.22 MB |  7.06x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.27x slower | 473000.0000 | 42000.0000 | 41000.0000 | 3578.89 MB | 11.22x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |   265.7 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.11x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.06x slower | 192000.0000 |  1000.0000 |          - | 1534.91 MB |  5.78x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.17x faster | 123000.0000 |  1000.0000 |          - |  983.37 MB |  3.70x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.19x slower | 207000.0000 |  1000.0000 |          - |  1656.2 MB |  6.23x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.10x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.08x slower | 195000.0000 |  1000.0000 |          - | 1558.22 MB |  5.30x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.20x faster | 126000.0000 |  1000.0000 |          - | 1007.03 MB |  3.42x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.35x slower | 212000.0000 |  1000.0000 |          - | 1694.93 MB |  5.76x more |
```

```
| Method      | ranger               | Mean     | Error     | StdDev   | Gen0         | Gen1        | Gen2       | Allocated |
|------------ |--------------------- |---------:|----------:|---------:|-------------:|------------:|-----------:|----------:|
| Access100mb | Excel(...)Cells [41] |  7.662 s |  0.1705 s | 0.0093 s |  175000.0000 |  98000.0000 |  7000.0000 |   1.49 GB |
| Access100mb | Excel(...)edXML [39] | 41.662 s | 10.9401 s | 0.5997 s | 1248000.0000 | 382000.0000 |  6000.0000 |  10.21 GB |
| Access100mb | Excel(...)PPlus [36] | 11.506 s |  2.5959 s | 0.1423 s |  550000.0000 | 188000.0000 | 13000.0000 |   6.08 GB |
| Access100mb | Excel(...)Prime [40] |  6.948 s |  0.2571 s | 0.0141 s |  204000.0000 |  58000.0000 |  6000.0000 |   1.56 GB |
| Access100mb | Excel(...)Spire [39] | 23.756 s |  1.3510 s | 0.0741 s | 1181000.0000 | 449000.0000 |  6000.0000 |   9.95 GB |
```
-----

# 2025-12-02 - Beta V2
- Remove explicit GC.Collect calls
- Replace per-iteration List<> with array (Cannot return sharedPool due to `yield return`)

```
| Method                          | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.08x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.59 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.54x slower | 387000.0000 | 52000.0000 |  3000.0000 | 3073.08 MB |  9.07x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.28x slower | 292000.0000 | 53000.0000 |  6000.0000 | 2291.71 MB |  6.77x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.53x slower | 387000.0000 | 52000.0000 |  3000.0000 | 3074.83 MB |  9.08x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.02x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.22x slower | 433000.0000 | 42000.0000 | 41000.0000 | 3260.61 MB | 10.22x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.06x faster | 282000.0000 | 42000.0000 | 41000.0000 | 2055.34 MB |  6.45x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.29x slower | 449000.0000 | 42000.0000 | 41000.0000 |    3384 MB | 10.61x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.11x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.04x slower | 180000.0000 |  1000.0000 |          - | 1439.61 MB |  5.42x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.19x faster | 111000.0000 |  1000.0000 |          - |     888 MB |  3.34x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.17x slower | 195000.0000 |  1000.0000 |          - | 1560.79 MB |  5.87x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.08x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.09x slower | 183000.0000 |  1000.0000 |          - | 1462.85 MB |  4.97x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.19x faster | 114000.0000 |  1000.0000 |          - |  911.66 MB |  3.10x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.37x slower | 200000.0000 |  1000.0000 |          - | 1599.56 MB |  5.44x more |
```
-----

# 2025-12-04 - Beta V2
- Removed finalizers from several classes
- Added a lightweight Row pooling strategy
```
| Method                          | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.73 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.07x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.53x slower | 375000.0000 | 56000.0000 |  6000.0000 | 2953.76 MB |  8.72x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.26x slower | 277000.0000 | 53000.0000 |  6000.0000 |  2172.4 MB |  6.41x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.54x slower | 375000.0000 | 56000.0000 |  6000.0000 |  2955.5 MB |  8.73x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.01x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.23x slower | 420000.0000 | 42000.0000 | 41000.0000 | 3151.48 MB |  9.88x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.09x faster | 269000.0000 | 42000.0000 | 41000.0000 |  1946.2 MB |  6.10x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.27x slower | 435000.0000 | 42000.0000 | 41000.0000 | 3274.87 MB | 10.27x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.09x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.04x slower | 173000.0000 |  1000.0000 |          - | 1386.21 MB |  5.22x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.20x faster | 104000.0000 |  1000.0000 |          - |  834.59 MB |  3.14x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.18x slower | 188000.0000 |  1000.0000 |          - | 1507.42 MB |  5.67x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.12x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.07x slower | 176000.0000 |  1000.0000 |          - | 1409.45 MB |  4.79x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.21x faster | 107000.0000 |  1000.0000 |          - |  858.25 MB |  2.92x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.35x slower | 193000.0000 |  1000.0000 |          - | 1546.16 MB |  5.26x more |
```
-----

# 2025-12-07 - Beta V2
- Investigate _memory usage_(s) 🧑‍💻
    - Removed finalizers from several classes
    - Added a lightweight Row pooling strategy
    - Replaced the `ReadString` implementation
    - Prefer returning `ReadOnlyMemory<char>` to avoid allocating a new string
    - Replace uses of `string.Split` in DefinedRange with span-based parsing
- Some code Cleanup
- Sacrificed a little speed..

```
| Method                          | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 3.95x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.40x slower | 282000.0000 | 46000.0000 |  6000.0000 |  2218.3 MB |  6.55x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.17x slower | 184000.0000 | 48000.0000 |  6000.0000 | 1436.91 MB |  4.24x more |
| NumberCellAsyncExcel_Prime      | 100mb.xlsx           | 1.41x slower | 283000.0000 | 45000.0000 |  6000.0000 | 2220.04 MB |  6.55x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.01x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.23x slower | 381000.0000 | 42000.0000 | 41000.0000 | 2841.84 MB |  8.91x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.06x faster | 230000.0000 | 42000.0000 | 41000.0000 | 1636.49 MB |  5.13x more |
| NumberCellAsyncExcel_Prime      | Blank(...).xlsx [30] | 1.28x slower | 396000.0000 | 42000.0000 | 41000.0000 | 2965.24 MB |  9.30x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.12x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.05x slower | 164000.0000 |  1000.0000 |          - |  1309.8 MB |  4.93x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.21x faster |  95000.0000 |  1000.0000 |          - |  758.17 MB |  2.85x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [34] | 1.17x slower | 179000.0000 |  1000.0000 |          - | 1430.96 MB |  5.39x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.10x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.10x slower | 167000.0000 |  1000.0000 |          - | 1332.98 MB |  4.53x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.18x faster |  98000.0000 |  1000.0000 |          - |  781.82 MB |  2.66x more |
| NumberCellAsyncExcel_Prime      | sampl(...).xlsx [30] | 1.38x slower | 184000.0000 |  1000.0000 |          - | 1469.73 MB |  5.00x more |
```
-----

# 2025-12-10
- User defined, using the `"A1:B10"` or `"$A$1:$B$10"` syntax

```
.NET SDK 10.0.100
  [Host]   : .NET 8.0.22 (8.0.22, 8.0.2225.52707), X64 RyuJIT x86-64-v3

| Method      | ranger               | Mean     | Error    | StdDev   | Gen0         | Gen1        | Gen2       | Allocated   |
|------------ |--------------------- |---------:|---------:|---------:|-------------:|------------:|-----------:|------------:|
| Access100mb | Excel(...)Cells [35] |  9.723 s |  5.482 s | 0.3005 s |  175000.0000 |  98000.0000 |  7000.0000 |  1527.47 MB |
| Access100mb | Excel(...)edXML [33] | 50.658 s | 50.775 s | 2.7831 s | 1248000.0000 | 382000.0000 |  6000.0000 | 10454.98 MB |
| Access100mb | Excel(...)PPlus [30] | 13.308 s |  2.008 s | 0.1101 s |  550000.0000 | 188000.0000 | 13000.0000 |  6227.98 MB |
| Access100mb | Excel(...)Prime [34] |  7.989 s |  2.612 s | 0.1432 s |  110000.0000 |  52000.0000 |  6000.0000 |   841.56 MB |
| Access100mb | Excel(...)Spire [33] | 29.134 s | 10.542 s | 0.5778 s | 1181000.0000 | 448000.0000 |  6000.0000 | 10192.05 MB |
```
-----

# 2025-12-20
- 1st pass of the performance for Extracting every Cell in XLSB

```
| Method                              | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------------------ |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan               | 100mb.xlsb           |     baseline |  43000.0000 | 41000.0000 | 5000.0000 |  335.54 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | 100mb.xlsb           | 2.86x slower | 521000.0000 | 55000.0000 | 3000.0000 | 4148.76 MB | 12.36x more |
| AccessEveryCellExcel_PrimeXlsb      | 100mb.xlsb           | 2.05x slower | 260000.0000 | 49000.0000 | 3000.0000 | 2065.03 MB |  6.15x more |
| NumberCellAsyncExcel_PrimeXlsb      | 100mb.xlsb           | 2.89x slower | 521000.0000 | 55000.0000 | 3000.0000 | 4148.76 MB | 12.36x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | Blank(...).xlsb [30] |     baseline |  37000.0000 |  1000.0000 |         - |  301.88 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | Blank(...).xlsb [30] | 3.51x slower | 579000.0000 |          - |         - | 4622.77 MB | 15.31x more |
| AccessEveryCellExcel_PrimeXlsb      | Blank(...).xlsb [30] | 2.07x slower | 268000.0000 |  1000.0000 |         - | 2144.35 MB |  7.10x more |
| NumberCellAsyncExcel_PrimeXlsb      | Blank(...).xlsb [30] | 3.53x slower | 579000.0000 |  1000.0000 |         - | 4622.77 MB | 15.31x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | sampl(...).xlsb [34] |     baseline |  32000.0000 |          - |         - |  262.23 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | sampl(...).xlsb [34] | 2.87x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
| AccessEveryCellExcel_PrimeXlsb      | sampl(...).xlsb [34] | 1.87x slower | 144000.0000 |  1000.0000 |         - | 1154.95 MB |  4.40x more |
| NumberCellAsyncExcel_PrimeXlsb      | sampl(...).xlsb [34] | 2.94x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | sampl(...).xlsb [30] |     baseline |  32000.0000 |          - |         - |  262.23 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | sampl(...).xlsb [30] | 2.89x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
| AccessEveryCellExcel_PrimeXlsb      | sampl(...).xlsb [30] | 1.82x slower | 144000.0000 |  1000.0000 |         - | 1154.95 MB |  4.40x more |
| NumberCellAsyncExcel_PrimeXlsb      | sampl(...).xlsb [30] | 2.86x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
```
-----

# 2025-12-21
- Performance improvements:
  - Switched from traditional switch statement to expression-based switch for more efficient dispatch
  - Delayed cell object allocation until after record type validation, avoiding unnecessary object creation for invalid records
  - Use of `[MethodImpl(MethodImplOptions.AggressiveInlining)]`
  - Detect ASCII-only strings by checking if high byte of each UTF-16 LE code unit is zero
```
| Method                              | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------------------ |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan               | 100mb.xlsb           |     baseline |  43000.0000 | 41000.0000 | 5000.0000 |  335.54 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | 100mb.xlsb           | 2.51x slower | 524000.0000 | 58000.0000 | 6000.0000 | 4148.61 MB | 12.36x more |
| AccessEveryCellExcel_PrimeXlsb      | 100mb.xlsb           | 1.69x slower | 260000.0000 | 49000.0000 | 3000.0000 | 2064.87 MB |  6.15x more |
| NumberCellAsyncExcel_PrimeXlsb      | 100mb.xlsb           | 2.53x slower | 524000.0000 | 58000.0000 | 6000.0000 | 4148.61 MB | 12.36x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | Blank(...).xlsb [30] |     baseline |  37000.0000 |  1000.0000 |         - |  301.88 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | Blank(...).xlsb [30] | 3.77x slower | 579000.0000 |  1000.0000 |         - | 4622.46 MB | 15.31x more |
| AccessEveryCellExcel_PrimeXlsb      | Blank(...).xlsb [30] | 2.01x slower | 268000.0000 |  1000.0000 |         - | 2144.04 MB |  7.10x more |
| NumberCellAsyncExcel_PrimeXlsb      | Blank(...).xlsb [30] | 3.60x slower | 579000.0000 |  1000.0000 |         - | 4622.46 MB | 15.31x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | sampl(...).xlsb [34] |     baseline |  32000.0000 |          - |         - |  262.23 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | sampl(...).xlsb [34] | 2.90x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
| AccessEveryCellExcel_PrimeXlsb      | sampl(...).xlsb [34] | 1.79x slower | 144000.0000 |  1000.0000 |         - | 1154.95 MB |  4.40x more |
| NumberCellAsyncExcel_PrimeXlsb      | sampl(...).xlsb [34] | 2.86x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | sampl(...).xlsb [30] |     baseline |  32000.0000 |          - |         - |  262.23 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | sampl(...).xlsb [30] | 2.78x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
| AccessEveryCellExcel_PrimeXlsb      | sampl(...).xlsb [30] | 1.82x slower | 144000.0000 |  1000.0000 |         - | 1154.95 MB |  4.40x more |
| NumberCellAsyncExcel_PrimeXlsb      | sampl(...).xlsb [30] | 2.84x slower | 312000.0000 |  1000.0000 |         - | 2494.25 MB |  9.51x more |
```
-----

# 2026-01-02
- Removal of the Conversion options `Number###`
- Implement "On Demand" conversion
    - Slightly slower, but less memory pressure for `xslb`

```
| Method                          | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan           | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper       | 100mb.xlsx           | 4.30x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.59 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime | 100mb.xlsx           | 1.45x slower | 296000.0000 | 40000.0000 |  2000.0000 | 2355.13 MB |  6.95x more |
| AccessEveryCellExcel_Prime      | 100mb.xlsx           | 1.20x slower | 202000.0000 | 49000.0000 |  6000.0000 | 1573.42 MB |  4.65x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper       | Blank(...).xlsx [30] | 1.03x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime | Blank(...).xlsx [30] | 1.27x slower | 411000.0000 | 42000.0000 | 41000.0000 | 3078.16 MB |  9.65x more |
| AccessEveryCellExcel_Prime      | Blank(...).xlsx [30] | 1.00x faster | 259000.0000 | 42000.0000 | 41000.0000 | 1872.57 MB |  5.87x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [34] | 1.14x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [34] | 1.06x slower | 177000.0000 |  1000.0000 |          - | 1416.69 MB |  5.33x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [34] | 1.22x faster | 108000.0000 |  1000.0000 |          - |  865.04 MB |  3.26x more |
|                                 |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan           | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper       | sampl(...).xlsx [30] | 1.11x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime | sampl(...).xlsx [30] | 1.12x slower | 180000.0000 |  1000.0000 |          - |  1440.2 MB |  4.90x more |
| AccessEveryCellExcel_Prime      | sampl(...).xlsx [30] | 1.15x faster | 111000.0000 |  1000.0000 |          - |  888.69 MB |  3.02x more |
```
```
| Method                              | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|------------------------------------ |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| AccessEveryCellSylvan               | 100mb.xlsb           |     baseline |  43000.0000 | 41000.0000 | 5000.0000 |  335.55 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | 100mb.xlsb           | 2.74x slower | 541000.0000 | 61000.0000 | 6000.0000 |  4282.9 MB | 12.76x more |
| AccessEveryCellExcel_PrimeXlsb      | 100mb.xlsb           | 1.85x slower | 277000.0000 | 50000.0000 | 3000.0000 | 2199.17 MB |  6.55x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | Blank(...).xlsb [30] |     baseline |  37000.0000 |  1000.0000 |         - |  301.88 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | Blank(...).xlsb [30] | 3.69x slower | 594000.0000 |  1000.0000 |         - | 4739.09 MB | 15.70x more |
| AccessEveryCellExcel_PrimeXlsb      | Blank(...).xlsb [30] | 2.24x slower | 283000.0000 |  1000.0000 |         - | 2260.49 MB |  7.49x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | sampl(...).xlsb [34] |     baseline |  32000.0000 |          - |         - |  262.23 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | sampl(...).xlsb [34] | 2.89x slower | 313000.0000 |  1000.0000 |         - | 2498.07 MB |  9.53x more |
| AccessEveryCellExcel_PrimeXlsb      | sampl(...).xlsb [34] | 1.89x slower | 145000.0000 |  1000.0000 |         - | 1158.76 MB |  4.42x more |
|                                     |                      |              |             |            |           |            |             |
| AccessEveryCellSylvan               | sampl(...).xlsb [30] |     baseline |  32000.0000 |          - |         - |  262.23 MB |             |
| AccessEveryCellAsyncExcel_PrimeXlsb | sampl(...).xlsb [30] | 2.92x slower | 313000.0000 |  1000.0000 |         - | 2498.06 MB |  9.53x more |
| AccessEveryCellExcel_PrimeXlsb      | sampl(...).xlsb [30] | 1.91x slower | 145000.0000 |  1000.0000 |         - | 1158.76 MB |  4.42x more |
```
-----


# 2026-01-04
- Change `GetAllCells` to return `IReadOnlyList<ICell?>?`
- Make use of more `[MethodImpl(MethodImplOptions.AggressiveOptimization)]`
- More XLSB benchmarks
```
| Method                                 | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|--------------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan                  | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.74 MB |             |
| AccessEveryCellXlsxHelper              | 100mb.xlsx           | 4.04x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime        | 100mb.xlsx           | 1.35x slower | 296000.0000 | 44000.0000 |  6000.0000 | 2327.23 MB |  6.87x more |
| AccessEveryCellExcel_Prime             | 100mb.xlsx           | 1.15x slower | 199000.0000 | 48000.0000 |  6000.0000 | 1554.39 MB |  4.59x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.28x slower | 511000.0000 | 39000.0000 |  2000.0000 | 4060.23 MB | 11.99x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper              | Blank(...).xlsx [30] | 1.03x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime        | Blank(...).xlsx [30] | 1.17x slower | 409000.0000 | 42000.0000 | 41000.0000 | 3062.18 MB |  9.60x more |
| AccessEveryCellExcel_Prime             | Blank(...).xlsx [30] | 1.06x faster | 258000.0000 | 42000.0000 | 41000.0000 | 1856.87 MB |  5.82x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.52x slower | 817000.0000 | 84000.0000 | 83000.0000 | 6146.63 MB | 19.27x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |   265.7 MB |             |
| AccessEveryCellXlsxHelper              | sampl(...).xlsx [34] | 1.10x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime        | sampl(...).xlsx [34] | 1.03x slower | 176000.0000 |  1000.0000 |          - | 1408.88 MB |  5.30x more |
| AccessEveryCellExcel_Prime             | sampl(...).xlsx [34] | 1.18x faster | 107000.0000 |  1000.0000 |          - |  857.35 MB |  3.23x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.21x slower | 354000.0000 |  1000.0000 |          - | 2826.03 MB | 10.64x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper              | sampl(...).xlsx [30] | 1.08x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime        | sampl(...).xlsx [30] | 1.07x slower | 179000.0000 |  1000.0000 |          - | 1432.16 MB |  4.87x more |
| AccessEveryCellExcel_Prime             | sampl(...).xlsx [30] | 1.17x faster | 110000.0000 |  1000.0000 |          - |  881.01 MB |  2.99x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.24x slower | 359000.0000 |  1000.0000 |          - | 2870.21 MB |  9.76x more |
```
```
| Method                                     | FileName             | Ratio         | Gen0         | Gen1        | Gen2      | Allocated   | Alloc Ratio |
|------------------------------------------- |--------------------- |--------------:|-------------:|------------:|----------:|------------:|------------:|
| AccessEveryCellSylvan                      | 100mb.xlsb           |      baseline |   43000.0000 |  41000.0000 | 5000.0000 |   335.55 MB |             |
| AccessEveryCellAspose                      | 100mb.xlsb           |  6.46x slower | 1865000.0000 | 127000.0000 | 8000.0000 | 15001.43 MB | 44.71x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | 100mb.xlsb           |  2.47x slower |  538000.0000 |  60000.0000 | 6000.0000 |   4255.7 MB | 12.68x more |
| AccessEveryCellExcel_PrimeXlsb             | 100mb.xlsb           |  1.85x slower |  278000.0000 |  53000.0000 | 6000.0000 |   2180.4 MB |  6.50x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           |  3.92x slower |  977000.0000 |  58000.0000 | 6000.0000 |  7735.68 MB | 23.05x more |
|                                            |                      |               |              |             |           |             |             |
| AccessEveryCellSylvan                      | Blank(...).xlsb [30] |      baseline |   37000.0000 |   1000.0000 |         - |   301.88 MB |             |
| AccessEveryCellAspose                      | Blank(...).xlsb [30] | 13.37x slower | 2402000.0000 | 617000.0000 | 7000.0000 | 19173.75 MB | 63.51x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | Blank(...).xlsb [30] |  3.22x slower |  592000.0000 |   1000.0000 |         - |  4723.49 MB | 15.65x more |
| AccessEveryCellExcel_PrimeXlsb             | Blank(...).xlsb [30] |  2.12x slower |  281000.0000 |   1000.0000 |         - |  2244.92 MB |  7.44x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] |  6.71x slower | 1184000.0000 |           - |         - |     9449 MB | 31.30x more |
|                                            |                      |               |              |             |           |             |             |
| AccessEveryCellSylvan                      | sampl(...).xlsb [34] |      baseline |   32000.0000 |           - |         - |   262.23 MB |             |
| AccessEveryCellAspose                      | sampl(...).xlsb [34] |  5.96x slower |  947000.0000 |  62000.0000 | 5000.0000 |  7523.41 MB | 28.69x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | sampl(...).xlsb [34] |  2.70x slower |  312000.0000 |   1000.0000 |         - |  2490.42 MB |  9.50x more |
| AccessEveryCellExcel_PrimeXlsb             | sampl(...).xlsb [34] |  1.88x slower |  144000.0000 |   1000.0000 |         - |  1151.14 MB |  4.39x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] |  5.40x slower |  624000.0000 |   1000.0000 |         - |  4981.47 MB | 19.00x more |
|                                            |                      |               |              |             |           |             |             |
| AccessEveryCellSylvan                      | sampl(...).xlsb [30] |      baseline |   32000.0000 |           - |         - |   262.23 MB |             |
| AccessEveryCellAspose                      | sampl(...).xlsb [30] |  5.93x slower |  959000.0000 |  62000.0000 | 5000.0000 |  7621.49 MB | 29.06x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | sampl(...).xlsb [30] |  2.61x slower |  312000.0000 |   1000.0000 |         - |  2490.43 MB |  9.50x more |
| AccessEveryCellExcel_PrimeXlsb             | sampl(...).xlsb [30] |  1.89x slower |  144000.0000 |   1000.0000 |         - |  1151.14 MB |  4.39x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] |  5.42x slower |  624000.0000 |   1000.0000 |         - |   4981.4 MB | 19.00x more |
```
-----

# 2026-01-11
- Clear Mod
- ThreadStringBuilderPool initial Size
- Row ThreadStatic
- Cell Address caching
- Make use of more `[MethodImpl(MethodImplOptions.AggressiveOptimization)]`
```
| Method                                 | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|--------------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan                  | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| AccessEveryCellXlsxHelper              | 100mb.xlsx           | 4.27x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.59 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime        | 100mb.xlsx           | 1.36x slower | 262000.0000 | 43000.0000 |  6000.0000 | 2054.67 MB |  6.07x more |
| AccessEveryCellExcel_Prime             | 100mb.xlsx           | 1.16x slower | 165000.0000 | 47000.0000 |  6000.0000 | 1281.85 MB |  3.78x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.31x slower | 446000.0000 | 51000.0000 |  6000.0000 | 3515.36 MB | 10.38x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper              | Blank(...).xlsx [30] | 1.02x faster | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime        | Blank(...).xlsx [30] | 1.10x slower | 349000.0000 | 42000.0000 | 41000.0000 | 2590.25 MB |  8.12x more |
| AccessEveryCellExcel_Prime             | Blank(...).xlsx [30] | 1.12x faster | 198000.0000 | 42000.0000 | 41000.0000 | 1384.94 MB |  4.34x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.42x slower | 699000.0000 | 84000.0000 | 83000.0000 | 5202.67 MB | 16.31x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| AccessEveryCellXlsxHelper              | sampl(...).xlsx [34] | 1.10x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime        | sampl(...).xlsx [34] | 1.03x faster | 149000.0000 |  1000.0000 |          - | 1195.36 MB |  4.50x more |
| AccessEveryCellExcel_Prime             | sampl(...).xlsx [34] | 1.26x faster |  80000.0000 |  1000.0000 |          - |  643.72 MB |  2.42x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.10x slower | 300000.0000 |  1000.0000 |          - | 2398.88 MB |  9.03x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper              | sampl(...).xlsx [30] | 1.08x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime        | sampl(...).xlsx [30] | 1.04x slower | 152000.0000 |  1000.0000 |          - | 1218.58 MB |  4.14x more |
| AccessEveryCellExcel_Prime             | sampl(...).xlsx [30] | 1.22x faster |  83000.0000 |  1000.0000 |          - |  667.38 MB |  2.27x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.14x slower | 306000.0000 |  1000.0000 |          - | 2443.15 MB |  8.31x more |
```
```
| Method                                     | FileName             | Ratio         | Gen0         | Gen1        | Gen2      | Allocated   | Alloc Ratio |
|------------------------------------------- |--------------------- |--------------:|-------------:|------------:|----------:|------------:|------------:|
| AccessEveryCellSylvan                      | 100mb.xlsb           |      baseline |   43000.0000 |  41000.0000 | 5000.0000 |   335.54 MB |             |
| AccessEveryCellAspose                      | 100mb.xlsb           |  6.75x slower | 1865000.0000 | 127000.0000 | 8000.0000 | 15000.78 MB | 44.71x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | 100mb.xlsb           |  2.47x slower |  538000.0000 |  60000.0000 | 6000.0000 |   4255.7 MB | 12.68x more |
| AccessEveryCellExcel_PrimeXlsb             | 100mb.xlsb           |  1.83x slower |  278000.0000 |  53000.0000 | 6000.0000 |   2180.4 MB |  6.50x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           |  3.90x slower |  974000.0000 |  53000.0000 | 3000.0000 |  7735.69 MB | 23.05x more |
|                                            |                      |               |              |             |           |             |             |
| AccessEveryCellSylvan                      | Blank(...).xlsb [30] |      baseline |   37000.0000 |   1000.0000 |         - |   301.88 MB |             |
| AccessEveryCellAspose                      | Blank(...).xlsb [30] | 13.66x slower | 2402000.0000 | 617000.0000 | 7000.0000 |  19174.9 MB | 63.52x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | Blank(...).xlsb [30] |  3.16x slower |  592000.0000 |   1000.0000 |         - |  4723.49 MB | 15.65x more |
| AccessEveryCellExcel_PrimeXlsb             | Blank(...).xlsb [30] |  2.10x slower |  281000.0000 |   1000.0000 |         - |  2244.92 MB |  7.44x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] |  6.70x slower | 1184000.0000 |           - |         - |     9449 MB | 31.30x more |
|                                            |                      |               |              |             |           |             |             |
| AccessEveryCellSylvan                      | sampl(...).xlsb [34] |      baseline |   32000.0000 |           - |         - |   262.23 MB |             |
| AccessEveryCellAspose                      | sampl(...).xlsb [34] |  6.00x slower |  947000.0000 |  62000.0000 | 5000.0000 |  7523.44 MB | 28.69x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | sampl(...).xlsb [34] |  2.57x slower |  312000.0000 |   1000.0000 |         - |  2490.43 MB |  9.50x more |
| AccessEveryCellExcel_PrimeXlsb             | sampl(...).xlsb [34] |  1.81x slower |  144000.0000 |   1000.0000 |         - |  1151.14 MB |  4.39x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] |  5.27x slower |  624000.0000 |   1000.0000 |         - |  4981.41 MB | 19.00x more |
|                                            |                      |               |              |             |           |             |             |
| AccessEveryCellSylvan                      | sampl(...).xlsb [30] |      baseline |   32000.0000 |           - |         - |   262.23 MB |             |
| AccessEveryCellAspose                      | sampl(...).xlsb [30] |  5.87x slower |  959000.0000 |  62000.0000 | 5000.0000 |  7621.47 MB | 29.06x more |
| AccessEveryCellAsyncExcel_PrimeXlsb        | sampl(...).xlsb [30] |  2.62x slower |  312000.0000 |   1000.0000 |         - |  2490.43 MB |  9.50x more |
| AccessEveryCellExcel_PrimeXlsb             | sampl(...).xlsb [30] |  1.80x slower |  144000.0000 |   1000.0000 |         - |  1151.14 MB |  4.39x more |
| ParallelEveryCellAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] |  5.15x slower |  624000.0000 |   1000.0000 |         - |   4981.4 MB | 19.00x more |
```
-----

# 2026-01-16 - V3
- Remove some `AggressiveOptimization` and allow `i-cache` to do its job
- Implement "Hot-Paths" for cell type access
- Reduce some memory allocations for ReadOnly CellCollections
```
| Method                                 | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|--------------------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| AccessEveryCellSylvan                  | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.74 MB |             |
| AccessEveryCellXlsxHelper              | 100mb.xlsx           | 4.40x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AccessEveryCellAsyncExcel_Prime        | 100mb.xlsx           | 1.36x slower | 259000.0000 | 43000.0000 |  6000.0000 | 2029.18 MB |  5.99x more |
| AccessEveryCellExcel_Prime             | 100mb.xlsx           | 1.17x slower | 162000.0000 | 47000.0000 |  6000.0000 | 1256.28 MB |  3.71x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.28x slower | 440000.0000 | 50000.0000 |  6000.0000 | 3464.05 MB | 10.23x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| AccessEveryCellXlsxHelper              | Blank(...).xlsx [30] | 1.02x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AccessEveryCellAsyncExcel_Prime        | Blank(...).xlsx [30] | 1.11x slower | 347000.0000 | 42000.0000 | 41000.0000 | 2566.88 MB |  8.05x more |
| AccessEveryCellExcel_Prime             | Blank(...).xlsx [30] | 1.10x faster | 195000.0000 | 42000.0000 | 41000.0000 | 1361.55 MB |  4.27x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.39x slower | 693000.0000 | 83000.0000 | 83000.0000 | 5155.85 MB | 16.17x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |   265.7 MB |             |
| AccessEveryCellXlsxHelper              | sampl(...).xlsx [34] | 1.10x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AccessEveryCellAsyncExcel_Prime        | sampl(...).xlsx [34] | 1.02x faster | 148000.0000 |  1000.0000 |          - | 1183.96 MB |  4.46x more |
| AccessEveryCellExcel_Prime             | sampl(...).xlsx [34] | 1.24x faster |  79000.0000 |  1000.0000 |          - |  632.28 MB |  2.38x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.06x slower | 297000.0000 |  1000.0000 |          - | 2375.97 MB |  8.94x more |
|                                        |                      |              |             |            |            |            |             |
| AccessEveryCellSylvan                  | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| AccessEveryCellXlsxHelper              | sampl(...).xlsx [30] | 1.09x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AccessEveryCellAsyncExcel_Prime        | sampl(...).xlsx [30] | 1.01x slower | 151000.0000 |  1000.0000 |          - | 1207.21 MB |  4.10x more |
| AccessEveryCellExcel_Prime             | sampl(...).xlsx [30] | 1.22x faster |  82000.0000 |  1000.0000 |          - |  655.93 MB |  2.23x more |
| ParallelEveryCellAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.10x slower | 303000.0000 |  1000.0000 |          - | 2420.25 MB |  8.23x more |
```

# 2026-01-18 - V4 Alpha
- Internal implementation of `IOpenXmlWorkBookReader::GetSheetNames` now returns the relative path to the sheetName
```
| Method                  | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------ |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| Sylvan                  | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| XlsxHelper              | 100mb.xlsx           | 4.24x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.59 MB |  9.98x more |
| AsyncExcel_Prime        | 100mb.xlsx           | 1.39x slower | 259000.0000 | 43000.0000 |  6000.0000 | 2029.28 MB |  5.99x more |
| Excel_Prime             | 100mb.xlsx           | 1.19x slower | 162000.0000 | 47000.0000 |  6000.0000 | 1256.38 MB |  3.71x more |
| PrlAsyncExcel_PrimeTwice| 100mb.xlsx           | 2.27x slower | 440000.0000 | 50000.0000 |  6000.0000 |  3464.3 MB | 10.23x more |
|                         |                      |              |             |            |            |            |             |
| Sylvan                  | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| XlsxHelper              | Blank(...).xlsx [30] | 1.01x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime        | Blank(...).xlsx [30] | 1.08x slower | 347000.0000 | 42000.0000 | 41000.0000 | 2566.97 MB |  8.05x more |
| Excel_Prime             | Blank(...).xlsx [30] | 1.11x faster | 195000.0000 | 42000.0000 | 41000.0000 | 1361.65 MB |  4.27x more |
| PrlAsyncExcel_PrimeTwice| Blank(...).xlsx [30] | 2.40x slower | 693000.0000 | 85000.0000 | 83000.0000 | 5155.95 MB | 16.17x more |
|                         |                      |              |             |            |            |            |             |
| Sylvan                  | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper              | sampl(...).xlsx [34] | 1.12x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime        | sampl(...).xlsx [34] | 1.04x faster | 148000.0000 |  1000.0000 |          - | 1184.05 MB |  4.46x more |
| Excel_Prime             | sampl(...).xlsx [34] | 1.25x faster |  79000.0000 |  1000.0000 |          - |  632.38 MB |  2.38x more |
| PrlAsyncExcel_PrimeTwice| sampl(...).xlsx [34] | 2.05x slower | 297000.0000 |  1000.0000 |          - | 2376.06 MB |  8.94x more |
|                         |                      |              |             |            |            |            |             |
| Sylvan                  | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper              | sampl(...).xlsx [30] | 1.09x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime        | sampl(...).xlsx [30] | 1.03x slower | 151000.0000 |  1000.0000 |          - | 1207.31 MB |  4.10x more |
| Excel_Prime             | sampl(...).xlsx [30] | 1.22x faster |  82000.0000 |  1000.0000 |          - |  656.03 MB |  2.23x more |
| PrlAsyncExcel_PrimeTwice| sampl(...).xlsx [30] | 2.17x slower | 303000.0000 |          - |          - |  2420.4 MB |  8.23x more |
```
-----

# 2026-01-20 - V4 Alpha
- Use `Task` and reduce memory allocations in some hot paths
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.73 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.07x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.34x slower | 249000.0000 | 43000.0000 |  6000.0000 | 1952.64 MB |  5.76x more |
| Excel_Prime              | 100mb.xlsx           | 1.14x slower | 162000.0000 | 47000.0000 |  6000.0000 | 1256.38 MB |  3.71x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.27x slower | 421000.0000 | 50000.0000 |  6000.0000 |  3312.2 MB |  9.78x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.02x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.13x slower | 338000.0000 | 42000.0000 | 41000.0000 | 2496.88 MB |  7.83x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.11x faster | 195000.0000 | 42000.0000 | 41000.0000 | 1361.65 MB |  4.27x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.44x slower | 676000.0000 | 83000.0000 | 83000.0000 | 5018.12 MB | 15.74x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.06x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.01x faster | 144000.0000 |          - |          - |  1149.7 MB |  4.33x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.23x faster |  79000.0000 |  1000.0000 |          - |  632.38 MB |  2.38x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.07x slower | 289000.0000 |  1000.0000 |          - | 2308.61 MB |  8.69x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.10x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.05x slower | 147000.0000 |  1000.0000 |          - | 1172.93 MB |  3.99x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.24x faster |  82000.0000 |  1000.0000 |          - |  656.03 MB |  2.23x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.10x slower | 294000.0000 |  1000.0000 |          - | 2352.57 MB |  8.00x more |
```
-----

# 2026-01-31 - V4 Alpha
- Fix fallout from making `CellValue` is now a `class`
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  43000.0000 | 41000.0000 |  5000.0000 |  338.72 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.18x slower | 424000.0000 |  5000.0000 |  2000.0000 | 3380.59 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.37x slower | 275000.0000 | 43000.0000 |  6000.0000 | 2157.06 MB |  6.37x more |
| Excel_Prime              | 100mb.xlsx           | 1.18x slower | 187000.0000 | 48000.0000 |  6000.0000 | 1460.79 MB |  4.31x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.26x slower | 472000.0000 | 48000.0000 |  6000.0000 | 3720.94 MB | 10.99x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  40000.0000 |  2000.0000 |  1000.0000 |  318.93 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.03x slower | 218000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.10x slower | 382000.0000 | 42000.0000 | 41000.0000 | 2850.86 MB |  8.94x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.05x faster | 240000.0000 | 42000.0000 | 41000.0000 |  1715.6 MB |  5.38x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.40x slower | 765000.0000 | 83000.0000 | 83000.0000 | 5726.15 MB | 17.95x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  33000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.09x faster | 100000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.02x faster | 164000.0000 |  1000.0000 |          - | 1309.98 MB |  4.93x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.18x faster |  99000.0000 |  1000.0000 |          - |   792.6 MB |  2.98x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.07x slower | 329000.0000 |  1000.0000 |          - | 2629.01 MB |  9.90x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  36000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.09x faster |  93000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.02x slower | 167000.0000 |  1000.0000 |          - | 1333.24 MB |  4.53x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.19x faster | 102000.0000 |  1000.0000 |          - |  816.25 MB |  2.77x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.11x slower | 335000.0000 |  1000.0000 |          - | 2673.03 MB |  9.09x more |
```
-----

# 2026-04-05 - V4 Beta
- Change the Cell type back to a `record`
- Add Formatting Style tests
- Update testing nugets
- Mark as Beta

```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  29000.0000 | 27000.0000 |  4000.0000 |  338.72 MB |             |
| XlsxHelper               | 100mb.xlsx           | 3.67x slower | 282000.0000 |  2000.0000 |  1000.0000 | 3380.58 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.32x slower | 184000.0000 | 36000.0000 |  5000.0000 | 2157.12 MB |  6.37x more |
| Excel_Prime              | 100mb.xlsx           | 1.17x slower | 126000.0000 | 34000.0000 |  5000.0000 | 1460.77 MB |  4.31x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.19x slower | 315000.0000 | 42000.0000 |  5000.0000 | 3720.82 MB | 10.99x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  27000.0000 |  2000.0000 |  1000.0000 |  318.92 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.00x faster | 145000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.11x slower | 268000.0000 | 42000.0000 | 41000.0000 | 2850.86 MB |  8.94x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.04x faster | 173000.0000 | 42000.0000 | 41000.0000 | 1715.59 MB |  5.38x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.53x slower | 538000.0000 | 83000.0000 | 83000.0000 | 5726.03 MB | 17.95x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  22000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.08x faster |  66000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.02x faster | 109000.0000 |  1000.0000 |          - | 1309.97 MB |  4.93x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.14x faster |  66000.0000 |  1000.0000 |          - |  792.59 MB |  2.98x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.00x slower | 219000.0000 |  1000.0000 |          - | 2628.99 MB |  9.90x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  24000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.05x faster |  62000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.01x slower | 111000.0000 |  1000.0000 |          - | 1333.28 MB |  4.53x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.15x faster |  68000.0000 |  1000.0000 |          - |  816.25 MB |  2.77x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 1.99x slower | 223000.0000 |  1000.0000 |          - | 2673.04 MB |  9.09x more |
```
```
| Method                       | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|----------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                    | 100mb.xlsb           |     baseline |  29000.0000 | 28000.0000 | 4000.0000 |  335.55 MB |             |
| AsyncExcel_PrimeXlsb         | 100mb.xlsb           | 2.22x slower | 360000.0000 | 42000.0000 | 5000.0000 | 4264.26 MB | 12.71x more |
| Excel_PrimeXlsb              | 100mb.xlsb           | 1.68x slower | 193000.0000 | 43000.0000 | 5000.0000 | 2265.57 MB |  6.75x more |
| PrlAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           | 3.51x slower | 655000.0000 | 48000.0000 | 6000.0000 | 7752.78 MB | 23.10x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | Blank(...).xlsb [30] |     baseline |  25000.0000 |  1000.0000 |         - |  301.88 MB |             |
| AsyncExcel_PrimeXlsb         | Blank(...).xlsb [30] | 3.17x slower | 403000.0000 |  1000.0000 |         - | 4828.05 MB | 15.99x more |
| Excel_PrimeXlsb              | Blank(...).xlsb [30] | 2.17x slower | 202000.0000 |  1000.0000 |         - | 2419.62 MB |  8.02x more |
| PrlAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] | 6.83x slower | 807000.0000 |  1000.0000 |         - | 9658.44 MB | 31.99x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [34] |     baseline |  21000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [34] | 2.53x slower | 212000.0000 |  1000.0000 |         - | 2540.16 MB |  9.69x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [34] | 1.74x slower | 103000.0000 |  1000.0000 |         - | 1235.15 MB |  4.71x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] | 5.16x slower | 424000.0000 |  1000.0000 |         - |  5080.9 MB | 19.38x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [30] |     baseline |  21000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [30] | 2.52x slower | 212000.0000 |  1000.0000 |         - | 2540.16 MB |  9.69x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [30] | 1.71x slower | 103000.0000 |  1000.0000 |         - | 1235.15 MB |  4.71x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] | 5.04x slower | 424000.0000 |  1000.0000 |         - |  5080.9 MB | 19.38x more |
```

-----

# 2026-04-10 - V4 - Beta
- Remove `in` usages (Supposed to not benefit !)
- Make use of ThreadStatic in XlsbRow
- Remove secoundary usage of a struct
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  29000.0000 | 27000.0000 |  4000.0000 |  338.72 MB |             |
| XlsxHelper               | 100mb.xlsx           | 3.60x slower | 282000.0000 |  2000.0000 |  1000.0000 | 3380.58 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.30x slower | 184000.0000 | 36000.0000 |  5000.0000 | 2157.11 MB |  6.37x more |
| Excel_Prime              | 100mb.xlsx           | 1.16x slower | 126000.0000 | 34000.0000 |  5000.0000 | 1460.77 MB |  4.31x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.08x slower | 315000.0000 | 42000.0000 |  5000.0000 | 3720.81 MB | 10.98x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  27000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.01x slower | 145000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.12x slower | 268000.0000 | 42000.0000 | 41000.0000 | 2850.85 MB |  8.94x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.05x faster | 173000.0000 | 42000.0000 | 41000.0000 | 1715.59 MB |  5.38x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.52x slower | 538000.0000 | 83000.0000 | 83000.0000 | 5726.04 MB | 17.96x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  22000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.08x faster |  66000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.02x faster | 109000.0000 |  1000.0000 |          - | 1309.96 MB |  4.93x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.17x faster |  66000.0000 |  1000.0000 |          - |  792.59 MB |  2.98x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.35x slower | 219000.0000 |  1000.0000 |          - | 2628.94 MB |  9.90x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  24000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.07x faster |  62000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.02x slower | 111000.0000 |  1000.0000 |          - | 1333.23 MB |  4.53x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.17x faster |  68000.0000 |  1000.0000 |          - |  816.25 MB |  2.77x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.06x slower | 223000.0000 |  1000.0000 |          - | 2672.97 MB |  9.09x more |
```
```
| Method                       | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|----------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                    | 100mb.xlsb           |     baseline |  29000.0000 | 28000.0000 | 4000.0000 |  335.55 MB |             |
| AsyncExcel_PrimeXlsb         | 100mb.xlsb           | 2.27x slower | 360000.0000 | 42000.0000 | 5000.0000 | 4264.26 MB | 12.71x more |
| Excel_PrimeXlsb              | 100mb.xlsb           | 1.66x slower | 193000.0000 | 43000.0000 | 5000.0000 | 2265.57 MB |  6.75x more |
| PrlAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           | 3.57x slower | 654000.0000 | 51000.0000 | 6000.0000 | 7752.75 MB | 23.10x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | Blank(...).xlsb [30] |     baseline |  25000.0000 |  1000.0000 |         - |  301.88 MB |             |
| AsyncExcel_PrimeXlsb         | Blank(...).xlsb [30] | 3.22x slower | 403000.0000 |  1000.0000 |         - | 4828.05 MB | 15.99x more |
| Excel_PrimeXlsb              | Blank(...).xlsb [30] | 2.07x slower | 202000.0000 |  1000.0000 |         - | 2419.62 MB |  8.02x more |
| PrlAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] | 6.81x slower | 807000.0000 |  1000.0000 |         - | 9658.44 MB | 31.99x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [34] |     baseline |  21000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [34] | 2.49x slower | 212000.0000 |  1000.0000 |         - | 2540.16 MB |  9.69x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [34] | 1.72x slower | 103000.0000 |  1000.0000 |         - | 1235.15 MB |  4.71x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] | 5.06x slower | 424000.0000 |  1000.0000 |         - |  5080.9 MB | 19.38x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [30] |     baseline |  21000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [30] | 2.51x slower | 212000.0000 |  1000.0000 |         - | 2540.16 MB |  9.69x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [30] | 1.71x slower | 103000.0000 |  1000.0000 |         - | 1235.15 MB |  4.71x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] | 5.09x slower | 424000.0000 |  1000.0000 |         - | 5080.97 MB | 19.38x more |
```

# 2026-04-22 - V4 - RC2
- Implement `System.DBNull` return option, for empty cells
- Mark up usage of userdefined cells styles for V5
- - Update "Sylvan.Data.Excel" to Version="0.5.5"
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  29000.0000 | 27000.0000 |  4000.0000 |  338.72 MB |             |
| XlsxHelper               | 100mb.xlsx           | 5.00x slower | 282000.0000 |  6000.0000 |  2000.0000 | 3380.58 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.36x slower | 184000.0000 | 36000.0000 |  5000.0000 | 2156.93 MB |  6.37x more |
| Excel_Prime              | 100mb.xlsx           | 1.23x slower | 126000.0000 | 34000.0000 |  5000.0000 | 1460.77 MB |  4.31x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.24x slower | 315000.0000 | 42000.0000 |  5000.0000 |  3719.5 MB | 10.98x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  26000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.08x slower | 145000.0000 |  2000.0000 |  1000.0000 | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.07x slower | 268000.0000 | 43000.0000 | 42000.0000 | 2850.56 MB |  8.94x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.02x faster | 173000.0000 | 43000.0000 | 42000.0000 | 1715.59 MB |  5.38x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.46x slower | 538000.0000 | 86000.0000 | 83000.0000 | 5724.28 MB | 17.95x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  22000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.08x faster |  66000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.02x faster | 109000.0000 |  1000.0000 |          - | 1309.88 MB |  4.93x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.16x faster |  66000.0000 |  1000.0000 |          - |  792.59 MB |  2.98x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.15x slower | 219000.0000 |  2000.0000 |          - | 2628.26 MB |  9.89x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  24000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.04x faster |  62000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.08x slower | 111000.0000 |  1000.0000 |          - | 1333.13 MB |  4.53x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.06x faster |  68000.0000 |  1000.0000 |          - |  816.25 MB |  2.77x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.30x slower | 223000.0000 |  2000.0000 |          - | 2671.72 MB |  9.08x more |
```

# 2026-04-27 - V4 - RC3
- Release-specific optimizations added
  - EnableTrimAnalyzer: true
  - TieredCompilation: true
  - TieredCompilationQuickJit: true
  - TieredCompilationQuickJitForLoops: true
- Implement `System.DBNull` return option, for empty cells
  - Implement `INullRow` return option, for empty rows
  - Update tests to use `INullRow` detection
- Implement `GetCell###(string columnLetters, ...)` #8
- `ArrayPool` support has been added to ThreadStringBuilderPool using ArrayPool<char>.
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  29000.0000 | 27000.0000 |  4000.0000 |  338.72 MB |             |
| XlsxHelper               | 100mb.xlsx           | 3.60x slower | 282000.0000 |  2000.0000 |  1000.0000 | 3380.58 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.28x slower | 184000.0000 | 36000.0000 |  5000.0000 | 2151.64 MB |  6.35x more |
| Excel_Prime              | 100mb.xlsx           | 1.15x slower | 126000.0000 | 34000.0000 |  5000.0000 | 1459.35 MB |  4.31x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.14x slower | 314000.0000 | 42000.0000 |  5000.0000 | 3709.95 MB | 10.95x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  27000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.05x slower | 145000.0000 |  1000.0000 |          - | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.13x slower | 268000.0000 | 42000.0000 | 41000.0000 | 2850.83 MB |  8.94x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.01x slower | 173000.0000 | 42000.0000 | 41000.0000 | 1715.59 MB |  5.38x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.56x slower | 538000.0000 | 83000.0000 | 83000.0000 |  5726.2 MB | 17.96x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  22000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.09x faster |  66000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.00x slower | 109000.0000 |  1000.0000 |          - | 1309.97 MB |  4.93x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.17x faster |  66000.0000 |  1000.0000 |          - |  792.59 MB |  2.98x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.18x slower | 219000.0000 |  1000.0000 |          - | 2629.08 MB |  9.90x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  24000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.08x faster |  62000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.02x slower | 111000.0000 |  1000.0000 |          - | 1333.29 MB |  4.53x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.16x faster |  68000.0000 |  1000.0000 |          - |  816.25 MB |  2.77x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.09x slower | 223000.0000 |  1000.0000 |          - | 2672.97 MB |  9.09x more |
```

# 2026-05-05 - V4
- Small tweaks for performance
  - Remove `<PublishReadyToRun>true</PublishReadyToRun>`
  - Supports zero-allocation formatting on .NET 8+ via ISpanFormattable.
  - Cache allocation of the internal buffer for the StringBuilder in ThreadStringBuilderPool, to avoid the first allocation on each thread.
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  29000.0000 | 27000.0000 |  4000.0000 |  338.72 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.80x slower | 282000.0000 |  4000.0000 |  1000.0000 | 3380.58 MB |  9.98x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.31x slower | 184000.0000 | 36000.0000 |  5000.0000 | 2151.38 MB |  6.35x more |
| Excel_Prime              | 100mb.xlsx           | 1.28x slower | 126000.0000 | 34000.0000 |  5000.0000 | 1459.35 MB |  4.31x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.36x slower | 314000.0000 | 42000.0000 |  5000.0000 | 3706.94 MB | 10.94x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  26000.0000 |  2000.0000 |  1000.0000 |   318.9 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.05x slower | 145000.0000 |  2000.0000 |  1000.0000 | 1739.24 MB |  5.45x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.10x slower | 268000.0000 | 43000.0000 | 42000.0000 | 2850.63 MB |  8.94x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.00x slower | 173000.0000 | 43000.0000 | 42000.0000 | 1715.59 MB |  5.38x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.33x slower | 537000.0000 | 85000.0000 | 83000.0000 | 5721.11 MB | 17.94x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  22000.0000 |          - |          - |  265.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.08x faster |  66000.0000 |          - |          - |  799.73 MB |  3.01x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.13x faster | 109000.0000 |  1000.0000 |          - | 1309.89 MB |  4.93x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.13x faster |  66000.0000 |  1000.0000 |          - |  792.59 MB |  2.98x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.10x slower | 219000.0000 |  2000.0000 |          - | 2626.55 MB |  9.89x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  24000.0000 |          - |          - |  294.16 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.01x faster |  62000.0000 |          - |          - |  742.13 MB |  2.52x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.09x slower | 111000.0000 |  1000.0000 |          - |  1333.1 MB |  4.53x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.07x faster |  68000.0000 |  1000.0000 |          - |  816.25 MB |  2.77x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.10x slower | 223000.0000 |  1000.0000 |          - |  2671.4 MB |  9.08x more |
```

# 2026-05-12 - V4 - Bug fixes
- Return null when `EndValue` is spotted in the xml #22
- Return null for not found SheetId #23
- Update nugets
- Changed CellValue to an abstract base class and removed all [FieldOffset(0)] fields.
- Introduced internal sealed class CellValue<T> : CellValue to store values of type T without boxing for value types.
- Notes: 
  - Sylvan (base) has been updated and uses less memory now, and is also slightly faster :wink:
  - But the memory footprint of Prime is also lower caprared to the above :grinning:
```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  33000.0000 | 32000.0000 |  3000.0000 |  397.64 MB |             |
| XlsxHelper               | 100mb.xlsx           | 8.14x slower | 282000.0000 |  3000.0000 |  1000.0000 | 3380.58 MB |  8.50x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 3.82x slower | 178000.0000 | 36000.0000 |  5000.0000 | 2083.17 MB |  5.24x more |
| Excel_Prime              | 100mb.xlsx           | 3.50x slower | 120000.0000 | 34000.0000 |  5000.0000 | 1391.21 MB |  3.50x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 6.07x slower | 303000.0000 | 37000.0000 |  5000.0000 | 3571.93 MB |  8.98x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  20000.0000 |  3000.0000 |          - |  246.02 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.06x slower | 145000.0000 |  2000.0000 |  1000.0000 | 1739.24 MB |  7.07x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.15x slower | 258000.0000 | 43000.0000 | 42000.0000 | 2732.33 MB | 11.11x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.05x slower | 164000.0000 | 43000.0000 | 42000.0000 |  1597.6 MB |  6.49x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 1.30x slower | 518000.0000 | 83000.0000 | 83000.0000 | 5490.05 MB | 22.32x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  14000.0000 |  1000.0000 |          - |  178.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.04x faster |  66000.0000 |          - |          - |  799.73 MB |  4.48x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.03x slower | 105000.0000 |  1000.0000 |          - | 1256.55 MB |  7.03x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.13x faster |  61000.0000 |  1000.0000 |          - |  739.18 MB |  4.14x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.32x slower | 210000.0000 |  1000.0000 |          - | 2522.12 MB | 14.12x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  15000.0000 |  1000.0000 |          - |  185.72 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.05x faster |  62000.0000 |          - |          - |  742.13 MB |  4.00x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.06x slower | 106000.0000 |  1000.0000 |          - | 1279.79 MB |  6.89x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.09x faster |  63000.0000 |  1000.0000 |          - |  762.84 MB |  4.11x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.17x slower | 214000.0000 |  1000.0000 |          - |  2566.1 MB | 13.82x more |
```

# 2026-05-25 - V5 - Beta
- ✅ Cell object type 📅
    - ✅ Store cell _style_ type (see Options enum)
    - ✅ Unit Tests
    - ✅ Implement reading of the styles to determine the default `DateTime` / `DateOnly` / `TimeOnly` formats #19
**Code-Level Optimizations**
   - ✅ Implement `ISpanFormattable` in CellValue
   - ✅ Optimisationm in the Xlsb workflow
   - ✅ Return to the usage of the FieldOffsets to store the BCL type to prevent boxings in the hot paths
   - ✅ Usage of the Fast convertors *i.e. ToDecimal is 3 times faster than Convert.ToDecimal* #20
**Advanced Scenarios**
   - ✅ Enable `PublishTrimmed=true` with trim warnings resolved
   - ✅ Native AOT compilation testing
**Bug Fixes**
- Implement reading of the styles to determine the default `DateTime` / `DateOnly` / `TimeOnly` formats #19
- `AsDecimal` method has an issue where it produces incorrect precision, but only with default options #20
- When Attempting to use the "SkipRows" on a a sheet that has null rows to start with, causes infinite loop #27
- When opening the source file, then use "Sharing Mode" to allow it to be opened by other things! (i.e. 2 instances of this !) #28

```
| Method                   | FileName             | Ratio        | Gen0        | Gen1       | Gen2       | Allocated  | Alloc Ratio |
|------------------------- |--------------------- |-------------:|------------:|-----------:|-----------:|-----------:|------------:|
| SylvanRdr                | 100mb.xlsx           |     baseline |  33000.0000 | 32000.0000 |  3000.0000 |  397.64 MB |             |
| XlsxHelper               | 100mb.xlsx           | 4.01x slower | 282000.0000 |  2000.0000 |  1000.0000 | 3380.58 MB |  8.50x more |
| AsyncExcel_Prime         | 100mb.xlsx           | 1.42x slower | 189000.0000 | 36000.0000 |  5000.0000 | 2219.65 MB |  5.58x more |
| Excel_Prime              | 100mb.xlsx           | 1.29x slower | 131000.0000 | 34000.0000 |  5000.0000 | 1527.48 MB |  3.84x more |
| PrlAsyncExcel_PrimeTwice | 100mb.xlsx           | 2.34x slower | 326000.0000 | 43000.0000 |  5000.0000 | 3846.15 MB |  9.67x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | Blank(...).xlsx [30] |     baseline |  20000.0000 |  3000.0000 |          - |  245.89 MB |             |
| XlsxHelper               | Blank(...).xlsx [30] | 1.04x slower | 145000.0000 |  1000.0000 |          - | 1739.24 MB |  7.07x more |
| AsyncExcel_Prime         | Blank(...).xlsx [30] | 1.16x slower | 278000.0000 | 42000.0000 | 41000.0000 | 2968.79 MB | 12.07x more |
| Excel_Prime              | Blank(...).xlsx [30] | 1.01x slower | 183000.0000 | 42000.0000 | 41000.0000 | 1833.57 MB |  7.46x more |
| PrlAsyncExcel_PrimeTwice | Blank(...).xlsx [30] | 2.64x slower | 558000.0000 | 83000.0000 | 83000.0000 | 5962.08 MB | 24.25x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [34] |     baseline |  14000.0000 |  1000.0000 |          - |  178.67 MB |             |
| XlsxHelper               | sampl(...).xlsx [34] | 1.09x faster |  66000.0000 |          - |          - |  799.73 MB |  4.48x more |
| AsyncExcel_Prime         | sampl(...).xlsx [34] | 1.03x slower | 113000.0000 |  1000.0000 |          - | 1363.36 MB |  7.63x more |
| Excel_Prime              | sampl(...).xlsx [34] | 1.15x faster |  70000.0000 |  1000.0000 |          - |     846 MB |  4.73x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [34] | 2.47x slower | 228000.0000 |  1000.0000 |          - | 2735.79 MB | 15.31x more |
|                          |                      |              |             |            |            |            |             |
| SylvanRdr                | sampl(...).xlsx [30] |     baseline |  15000.0000 |  1000.0000 |          - |  185.74 MB |             |
| XlsxHelper               | sampl(...).xlsx [30] | 1.04x faster |  62000.0000 |          - |          - |  742.13 MB |  4.00x more |
| AsyncExcel_Prime         | sampl(...).xlsx [30] | 1.19x slower | 117000.0000 |  1000.0000 |          - | 1401.94 MB |  7.55x more |
| Excel_Prime              | sampl(...).xlsx [30] | 1.00x faster |  73000.0000 |  1000.0000 |          - |  884.91 MB |  4.76x more |
| PrlAsyncExcel_PrimeTwice | sampl(...).xlsx [30] | 2.27x slower | 234000.0000 |  1000.0000 |          - | 2810.31 MB | 15.13x more |
```

```
| Method                       | FileName             | Ratio        | Gen0        | Gen1       | Gen2      | Allocated  | Alloc Ratio |
|----------------------------- |--------------------- |-------------:|------------:|-----------:|----------:|-----------:|------------:|
| SylvanRdr                    | 100mb.xlsb           |     baseline |  29000.0000 | 28000.0000 | 4000.0000 |  335.55 MB |             |
| AsyncExcel_PrimeXlsb         | 100mb.xlsb           | 2.33x slower | 366000.0000 | 42000.0000 | 5000.0000 | 4332.39 MB | 12.91x more |
| Excel_PrimeXlsb              | 100mb.xlsb           | 1.65x slower | 199000.0000 | 43000.0000 | 5000.0000 |  2333.7 MB |  6.95x more |
| PrlAsyncExcel_PrimeXlsbTwice | 100mb.xlsb           | 3.79x slower | 666000.0000 | 49000.0000 | 6000.0000 | 7888.96 MB | 23.51x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | Blank(...).xlsb [30] |     baseline |  25000.0000 |  1000.0000 |         - |  301.88 MB |             |
| AsyncExcel_PrimeXlsb         | Blank(...).xlsb [30] | 3.32x slower | 413000.0000 |  1000.0000 |         - | 4946.03 MB | 16.38x more |
| Excel_PrimeXlsb              | Blank(...).xlsb [30] | 2.00x slower | 212000.0000 |  1000.0000 |         - | 2537.61 MB |  8.41x more |
| PrlAsyncExcel_PrimeXlsbTwice | Blank(...).xlsb [30] | 7.04x slower | 826000.0000 |  1000.0000 |         - | 9894.36 MB | 32.78x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [34] |     baseline |  21000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [34] | 2.47x slower | 216000.0000 |  1000.0000 |         - | 2590.55 MB |  9.88x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [34] | 1.58x slower | 107000.0000 |  1000.0000 |         - | 1285.55 MB |  4.90x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [34] | 4.92x slower | 433000.0000 |  1000.0000 |         - | 5181.67 MB | 19.76x more |
|                              |                      |              |             |            |           |            |             |
| SylvanRdr                    | sampl(...).xlsb [30] |     baseline |  21000.0000 |          - |         - |  262.23 MB |             |
| AsyncExcel_PrimeXlsb         | sampl(...).xlsb [30] | 2.51x slower | 216000.0000 |  1000.0000 |         - | 2590.57 MB |  9.88x more |
| Excel_PrimeXlsb              | sampl(...).xlsb [30] | 1.63x slower | 107000.0000 |  1000.0000 |         - | 1285.55 MB |  4.90x more |
| PrlAsyncExcel_PrimeXlsbTwice | sampl(...).xlsb [30] | 5.14x slower | 433000.0000 |  1000.0000 |         - | 5181.66 MB | 19.76x more |
```
```
| Method           | Mean     | Error    | Ratio | Allocated    | Alloc Ratio |
|----------------- |---------:|---------:|------:|-------------:|------------:|
| Baseline         | 109.7 ms | 12.51 ms |  1.00 |    246.66 KB |        1.00 |
| SylvanXlsx       | 153.2 ms | 13.59 ms |  1.40 |    665.48 KB |        2.70 |
| SylvanXlsxObj    | 167.6 ms | 27.50 ms |  1.53 |  14474.14 KB |       58.68 |
| SylvanXlsx_BindT | 176.5 ms | 48.98 ms |  1.61 |  11930.56 KB |       48.37 |
| PrimeXlsxObj     | 179.7 ms |  2.44 ms |  1.64 |    106499 KB |      431.77 |
| PrimeXlsxType    | 179.8 ms |  6.04 ms |  1.64 |  90174.36 KB |      365.59 |
| PrimeXlsx        | 194.2 ms |  2.56 ms |  1.77 | 112633.75 KB |      456.64 |
```