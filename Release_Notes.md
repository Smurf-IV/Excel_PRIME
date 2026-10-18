# 2026-10-18 - V5 nuget update
- Update nugets
- Made benchmarks run .net**9** and .net**10** as well
- **Memory & Allocation Optimizations**
- Some tinkering in the way memory is used and allocated, especially in the Xlsb workflow
- Eliminated String Caching Overhead: Removed the ConcurrentDictionary cache for small strings in CellValue, which introduced locking and hashing overhead during high-volume parsing.
- Simplified SIMD Decoding: Reverted SimdStringDecoder.cs to use Vector<ushort> instead of explicit Vector512/Vector256/Vector128 branching, reducing overhead for common short strings.
- Optimized Column Parsing: Replaced SearchValues with simple loops in ExcelColumns.cs for cell reference parsing, as the overhead of SearchValues was counterproductive for 1-3 character strings.
- Tuned Method Inlining: Removed AggressiveOptimization from several complex methods to allow better JIT tiered compilation and reduce instruction cache pressure.