## 2026-11-13 - V6 - Excel_FAST - Beta
- ✅ Add .Net11 TFM support
- Replaced the previous float round-trip memory comparison in CellValue.IsExactlyDouble with a fast memory-overlay check on the underlying 64-bit decimal mantissa (lo64 <= 999_999_999_999_999UL and hi == 0).

## 2026-11-12 - V6 - Excel_FAST - Alpha
- ⛓️‍💥 **Breaking Change(s)**
    - If you want to use `Async` then create a `new Excel_PRIMEAsync()`
    - Same applies to `Xlsb`
    - Use `ReadOnlySpan<Cell>` instead of `ArraySegment<Cell>`
      - Using foreach over ReadOnlySpan<T> is guaranteed to be allocation-free and boxing-free. If you need LINQ-style processing on a span without allocations, iterate using for or foreach directly over the span
- ✅ **ExcelFAST High-Throughput Subsystem**:
    - ✅ Dedicated zero-allocation sequential reader pipeline (`IExcelFast`, `ISheetFast`, `IRowFast`).
    - ✅ Format-specific implementations: `ExcelFastXlsx` for OpenXML (.xlsx) and `ExcelFastXlsb` for binary OpenXML (.xlsb).
    - ✅ Array pooling integration (`ReadOnlySpan<Cell>`) with `PooledCellView` and thread-static buffers for zero GC allocation during hot iteration loops.
    - ✅ Full support across .NET 8.0, .NET 9.0, and .NET 10.0.
    - ✅ Integrated test suite and benchmarking coverage in `ExcelFAST.Tests`.
- ✅ Performance Optimizations for Prime (See breaking changes above)

# 2026-10-18 - V5 nuget update
- Update nugets
- Made benchmarks run .net**9** and .net**10** as well
- **Memory & Allocation Optimizations**
- Some tinkering in the way memory is used and allocated, especially in the Xlsb workflow
- Eliminated String Caching Overhead: Removed the ConcurrentDictionary cache for small strings in CellValue, which introduced locking and hashing overhead during high-volume parsing.
- Simplified SIMD Decoding: Reverted SimdStringDecoder.cs to use Vector<ushort> instead of explicit Vector512/Vector256/Vector128 branching, reducing overhead for common short strings.
- Optimized Column Parsing: Replaced SearchValues with simple loops in ExcelColumns.cs for cell reference parsing, as the overhead of SearchValues was counterproductive for 1-3 character strings.
- Tuned Method Inlining: Removed AggressiveOptimization from several complex methods to allow better JIT tiered compilation and reduce instruction cache pressure.