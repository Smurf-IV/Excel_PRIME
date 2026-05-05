# Excel_PRIME Code-Level Performance Improvements - Implementation Plan

## Overview

This document outlines the next phase of performance optimizations for Excel_PRIME, focusing on code-level improvements that leverage .NET 8+ features for zero-allocation operations and SIMD optimizations.

## Current State Assessment

### Already Optimized
✅ **CellValue.cs**
- Struct layout optimization (LayoutKind.Explicit)
- Hot path separation (ToString fast path)
- MethodImpl.NoInlining for cold paths
- Cached CultureInfo for reduced allocations
- Switch expressions for early returns

✅ **Cell.cs**
- Record type for value semantics
- Column letter caching (static array)
- Async/await patterns
- Early exit conditions

✅ **Thread pools & Builders**
- ThreadStringBuilderPool for StringBuilder reuse
- Memory pooling patterns

## Recommended Improvements (Priority Order)

### Phase 1: Zero-Allocation Formatting (High Impact)

#### 1.1 Implement ISpanFormattable in CellValue
**Impact**: Eliminates string allocations in ToString operations

```csharp
// Before: String allocation on every ToString call
public override string ToString() => _doubleValue.ToString(InvariantCultureCache);

// After: Zero-allocation span-based formatting
public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
{
    return _type switch
    {
        CellValueType.Numeric => _doubleValue.TryFormat(destination, out charsWritten, format, provider),
        CellValueType.DateTime => _dateTimeValue.TryFormat(destination, out charsWritten, format, provider),
        CellValueType.Bool => _boolValue ? TryFormatTrue(destination, out charsWritten) : TryFormatFalse(destination, out charsWritten),
        _ => TryFormatNull(destination, out charsWritten)
    };
}
```

**Benefit**: 
- ~40-50% fewer allocations in format-heavy workloads
- Direct span writing, no intermediate strings
- Compatible with existing ToString() (acts as fallback)

**Files to Modify**:
- `CellValue.cs` - Add ISpanFormattable implementation

**Effort**: 2-3 hours

---

#### 1.2 Add ToString(StringBuilder) Overload
**Impact**: Reuse existing builders in hot paths

```csharp
internal void ToString(StringBuilder builder)
{
    if (_strValue != null || _type == CellValueType.String)
    {
        builder.Append(_strValue);
        return;
    }

    switch (_type)
    {
        case CellValueType.Numeric:
            builder.Append(_doubleValue);
            break;
        // ... other cases
    }
}
```

**Benefit**:
- Works with existing ThreadStringBuilderPool
- Zero additional allocations for frequently formatted values
- Direct integration with async reading pipeline

**Files to Modify**:
- `CellValue.cs` - Add StringBuilder overload
- `Cell.cs` - Use in row building

**Effort**: 1-2 hours

---

### Phase 2: Collection Optimizations (Medium Impact)

#### 2.1 Use CollectionsMarshal for Zero-Copy Operations
**Impact**: Eliminates intermediate array copies

**Current Pattern** (in various collection operations):
```csharp
// Creates temporary List and copies
var list = new List<ICell>(cells);
return list.AsReadOnly();
```

**Optimized Pattern**:
```csharp
// Direct span access without copying
public ReadOnlySpan<ICell> GetCells()
{
    return CollectionsMarshal.AsSpan(_cellList);
}
```

**Benefit**:
- 10-20% reduction in memory allocations for row/cell collections
- Direct span access for bulk operations
- Compatible with existing APIs via wrapper methods

**Files to Modify**:
- `Row.cs` - Cell collection access
- `Sheet.cs` - Row collection access
- `Implementation/ReaderAtoms.cs` - If using collections

**Effort**: 1-2 hours (search and replace safe)

---

#### 2.2 Add LINQ-to-Span Extensions
**Impact**: Efficient enumeration without allocation

```csharp
// In a new Extensions file or existing Extensions.cs
public static Span<T> AsSpan<T>(this List<T> list)
    => CollectionsMarshal.AsSpan(list);

public static bool Any<T>(this Span<T> span, Predicate<T> predicate)
{
    foreach (var item in span)
    {
        if (predicate(item)) return true;
    }
    return false;
}
```

**Benefit**:
- Enables zero-allocation LINQ-like operations
- Works with existing List collections
- Reduces foreach allocation patterns

**Files to Modify**:
- `FromExternal/Extensions.cs` - Add span extension methods

**Effort**: 2-3 hours

---

### Phase 3: SIMD Optimizations (Medium Impact)

#### 3.1 Add SIMD-Accelerated Cell Parsing
**Impact**: 2-4x faster for numeric bulk operations

**Scenario**: When reading large ranges of numeric cells

```csharp
// Using System.Runtime.Intrinsics for vectorized operations
internal static class SIMDHelper
{
    // Example: Parse multiple double values in parallel
    public static void ParseNumericBatch(ReadOnlySpan<string> values, Span<double> output)
    {
        int i;
        // Vector-based parsing for aligned data
        for (i = 0; i + Vector256<double>.Count <= values.Length; i += Vector256<double>.Count)
        {
            // SIMD operation on batch
            // Fallback to scalar for remainder
        }

        // Scalar fallback for remaining values
        for (; i < values.Length; i++)
        {
            double.TryParse(values[i], out output[i]);
        }
    }
}
```

**Benefit**:
- 2-4x faster numeric cell parsing in bulk operations
- Automatic fallback for non-aligned data
- Leverages modern CPU instruction sets

**Files to Create**:
- `Implementation/SIMDHelper.cs` - SIMD utilities

**Effort**: 3-4 hours (requires performance testing)

---

#### 3.2 Vectorized String Comparison
**Impact**: Faster column/row lookups

```csharp
// Fast string matching for cell references
internal static class StringVectorization
{
    public static int IndexOfOrdinal(ReadOnlySpan<char> text, ReadOnlySpan<char> pattern)
    {
        // Use Vector operations for pattern matching
        // Fallback to Span<T>.IndexOf for small patterns
        if (pattern.Length <= 4)
            return text.IndexOf(pattern);

        // SIMD-accelerated search for longer patterns
        return VectorizedSearch(text, pattern);
    }
}
```

**Benefit**:
- 3-5x faster cell reference lookups
- Especially beneficial for large range queries
- Works with existing column/row naming

**Files to Modify**:
- `FromExternal/ExcelColumns.cs` - Add vectorized lookup
- Create `Implementation/StringVectorization.cs` if needed

**Effort**: 2-3 hours

---

### Phase 4: Memory & Streaming Improvements (Lower Priority)

#### 4.1 Implement IAsyncEnumerable<IRow>
**Impact**: Memory-efficient streaming for large files

```csharp
public async IAsyncEnumerable<IRow> GetRowsAsync(
    [EnumeratorCancellation] CancellationToken ct = default)
{
    // Instead of loading all rows into memory
    foreach (var row in _rows) // Currently loads everything
    {
        yield return row;
    }
}
```

**Benefit**:
- Stream large files without loading into memory
- Compatible with async/await patterns
- Enables memory-constrained environments

**Files to Modify**:
- `ISheet.cs` - Add IAsyncEnumerable method
- `Implementation/Sheet.cs` - Implement streaming

**Effort**: 3-4 hours

---

#### 4.2 Add Pooled Memory<T> Support
**Impact**: Reduce GC pressure on large operations

```csharp
using var buffer = MemoryPool<char>.Shared.Rent(4096);
// Use buffer for formatting operations
```

**Benefit**:
- Reduces GC allocation pressure
- Enables large batch operations
- Works with existing thread pools

**Files to Modify**:
- `Implementation/ThreadStringBuilderPool.cs` - Add memory pool support

**Effort**: 2-3 hours

---

## Implementation Strategy

### Recommended Sequence

1. **Week 1-2**: Implement ISpanFormattable in CellValue (Phase 1.1)
   - Highest ROI
   - Affects core formatting path
   - Relatively isolated change

2. **Week 2-3**: Add StringBuilder overloads (Phase 1.2)
   - Integrates with existing infrastructure
   - Builds on Phase 1.1
   - Reduces allocations further

3. **Week 3-4**: CollectionsMarshal optimizations (Phase 2)
   - Safe, low-risk changes
   - Broad impact across collection operations
   - Good for iteration

4. **Week 4-5**: SIMD optimizations (Phase 3)
   - Requires more testing
   - High performance impact
   - Can be progressive (add individually)

5. **Week 5-6**: Streaming & memory (Phase 4)
   - Larger architectural changes
   - Optional for most use cases
   - Good for future versions

---

## Testing Strategy

### Performance Benchmarks to Add

```csharp
[Benchmark]
public string CellValue_ToString_Double()
{
    var cell = new CellValue(123.456, 0);
    return cell.ToString(); // Should show improvement
}

[Benchmark]
public void CellValue_TryFormat_Double()
{
    var cell = new CellValue(123.456, 0);
    Span<char> buffer = stackalloc char[32];
    cell.TryFormat(buffer, out _); // Should be zero-alloc
}

[Benchmark]
public void CollectionsMarshal_GetCells()
{
    var row = GetRow();
    var cells = CollectionsMarshal.AsSpan(row.Cells);
    // Verify no allocation
}

[Benchmark]
public void SIMD_ParseNumericBatch()
{
    var values = GetStringArray(1000);
    var output = new double[1000];
    SIMDHelper.ParseNumericBatch(values, output);
}
```

### Measurement Tools
- **BenchmarkDotNet** (already in project)
- **Allocation profiler** in Visual Studio
- **ETW profiling** for JIT metrics

---

## Risk Assessment

| Phase | Risk Level | Mitigation |
|-------|-----------|-----------|
| Phase 1 | Low | Small, isolated changes; good test coverage |
| Phase 2 | Low | Safe collection operations; no behavior change |
| Phase 3 | Medium | Requires careful fallback handling; platform-dependent |
| Phase 4 | Medium | Larger API changes; good for major version |

---

## Expected Performance Improvements

### Conservative Estimates
- **Phase 1** (ISpanFormattable): 10-15% reduction in allocations for text-heavy operations
- **Phase 2** (CollectionsMarshal): 5-10% fewer allocations in collection operations
- **Phase 3** (SIMD): 100-300% faster for bulk numeric operations
- **Phase 4** (Streaming): 50-90% less memory for large file processing

### Combined Impact
```
Small Files (< 1MB):      10-15% faster
Medium Files (1-100MB):   15-25% faster
Large Files (> 100MB):    25-40% faster (with streaming)
Batch Operations:         100-300% faster (with SIMD)
```

---

## Breaking Changes: NONE

All improvements are additive:
- Existing APIs preserved
- New methods added alongside existing ones
- ISpanFormattable automatically used by modern APIs
- Old ToString() still works

---

## Dependencies

All required types are in:
- `System.Runtime` (ISpanFormattable)
- `System.Collections` (CollectionsMarshal)
- `System.Runtime.Intrinsics` (SIMD) - Available .NET 8+

No new NuGet packages required.

---

## Success Criteria

- ✅ All benchmarks show improvement
- ✅ Zero new allocations in hot paths (Phase 1-2)
- ✅ SIMD implementation properly falls back on unsupported platforms
- ✅ All existing tests pass
- ✅ No breaking changes to public API
- ✅ Documentation updated

---

## Next Steps

1. **Review and approve** this plan
2. **Create GitHub issues** for each phase
3. **Set up performance baseline** with BenchmarkDotNet
4. **Implement Phase 1** (ISpanFormattable)
5. **Measure** before/after performance
6. **Iterate** through remaining phases

---

## References

- [ISpanFormattable Interface](https://learn.microsoft.com/en-us/dotnet/api/system.ispanformattable)
- [CollectionsMarshal Class](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.collectionsmarshal)
- [System.Runtime.Intrinsics](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.intrinsics)
- [.NET Performance Best Practices](https://learn.microsoft.com/en-us/dotnet/fundamentals/code-analysis/style-rules/performance-rules)

---

**Date**: 2024  
**Status**: Ready for Implementation  
**Target**: Excel_PRIME v4.2+  
