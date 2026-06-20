```

BenchmarkDotNet v0.15.8, Windows 11 (10.0.26200.8246/25H2/2025Update/HudsonValley2)
13th Gen Intel Core i5-13600KF 2.60GHz, 1 CPU, 20 logical and 14 physical cores
.NET SDK 10.0.203
  [Host]   : .NET 8.0.27 (8.0.27, 8.0.2726.22922), X64 RyuJIT x86-64-v3
  ShortRun : .NET 8.0.27 (8.0.27, 8.0.2726.22922), X64 RyuJIT x86-64-v3

Job=ShortRun  IterationCount=3  LaunchCount=1  
WarmupCount=3  

```
| Method   | Length | Mean      | Error     | StdDev    | Ratio         | RatioSD | Allocated | Alloc Ratio |
|--------- |------- |----------:|----------:|----------:|--------------:|--------:|----------:|------------:|
| **Slice**    | **10**     | **0.1840 ns** | **0.0061 ns** | **0.0003 ns** |      **baseline** |        **** |         **-** |          **NA** |
| Range    | 10     | 0.1849 ns | 0.0182 ns | 0.0010 ns |  1.01x slower |   0.00x |         - |          NA |
| SliceEnd | 10     | 0.0181 ns | 0.0832 ns | 0.0046 ns | 10.58x faster |   2.02x |         - |          NA |
| RangeEnd | 10     | 0.3602 ns | 0.1470 ns | 0.0081 ns |  1.96x slower |   0.04x |         - |          NA |
|          |        |           |           |           |               |         |           |             |
| **Slice**    | **100**    | **0.1844 ns** | **0.0692 ns** | **0.0038 ns** |      **baseline** |        **** |         **-** |          **NA** |
| Range    | 100    | 0.1818 ns | 0.0413 ns | 0.0023 ns |  1.01x faster |   0.02x |         - |          NA |
| SliceEnd | 100    | 0.0051 ns | 0.0225 ns | 0.0012 ns | 37.36x faster |   7.10x |         - |          NA |
| RangeEnd | 100    | 0.3679 ns | 0.1026 ns | 0.0056 ns |  2.00x slower |   0.04x |         - |          NA |
|          |        |           |           |           |               |         |           |             |
| **Slice**    | **1000**   | **0.1854 ns** | **0.0766 ns** | **0.0042 ns** |      **baseline** |        **** |         **-** |          **NA** |
| Range    | 1000   | 0.1831 ns | 0.0954 ns | 0.0052 ns |  1.01x faster |   0.03x |         - |          NA |
| SliceEnd | 1000   | 0.3167 ns | 6.2099 ns | 0.3404 ns |  1.71x slower |   1.59x |         - |          NA |
| RangeEnd | 1000   | 0.3457 ns | 0.0081 ns | 0.0004 ns |  1.87x slower |   0.04x |         - |          NA |
