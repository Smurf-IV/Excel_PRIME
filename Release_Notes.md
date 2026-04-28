# 2026-04-27 - V4 - RC3
- `ArrayPool` support has been added to ThreadStringBuilderPool using ArrayPool<char>.
- Release-specific optimizations added
  - EnableTrimAnalyzer: true
  - TieredCompilation: true
  - TieredCompilationQuickJit: true
  - TieredCompilationQuickJitForLoops: true
- Implement `System.DBNull` return option, for empty cells
  - Implement `INullRow` return option, for empty rows
  - Update tests to use `INullRow` detection
- Implement `GetCell###(string columnLetters, ...)` #8

# 2026-04-22 - V4 - RC2
- Implement `System.DBNull` return option, for empty cells
- Mark up usage of userdefined cells styles for V5
- Update "Sylvan.Data.Excel" to Version="0.5.5"
    - 🚀 [2026-04-22 - V4 - RC2](https://github.com/Smurf-IV/Excel_PRIME/blob/V4-Dev/Performance.md#2026-04-22---v4---rc2)

# 2026-04-16 - V4 - RC1
- Update nugets and update performance numbers
- 🚀 [2026-04-10](https://github.com/Smurf-IV/Excel_PRIME/blob/V4-Dev/Performance.md#2026-04-10---v4---beta)

# 2026-04-10 - V4 - Beta
- Remove `in` usages (Supposed to not benefit !)
- Make use of ThreadStatic in XlsbRow
- Remove secoundary usage of a struct

# 2026-04-05 - V4 - Beta
- ⛓️‍💥 **Breaking Change(s)**
    - Removal of `GetSheetFileName(int offsetSheetId);`
    - Removal of `GetDefinedRange` via `int sheetId`
    - Removal of `Index` property from `ISheet`
    - Internal Creation of WorkBooks
    - Internal implementation of `IOpenXmlWorkBookReader::GetSheetNames` now returns the relative path to the sheetName
    - `CellValue` is now a `class`, therefore no need to use `.Value`
    - `ICell.CellValue` is now nullable
- Re-introduce the CellConversion for the XLSX cell types
- Use `ValueTask` and reduce memory allocations in some hot paths
- Cell object type 📅
    - "Best Effort" `Operator` based conversion
    - TryGet`Type` will return `out type`, if stored as that type.
- Add `Ecma376StandardProvider`
- Add `StylesExtractor`
- Attempt to make use of the Cell types
- Tinker with some `MethodImpl`
- Add `_iStyleRef` and start to add formatting based on it
- Fix fallout from making `CellValue` is now a `class`
  - 🚀 [2026-01-31](https://github.com/Smurf-IV/Excel_PRIME/blob/V4-Dev/Performance.md#2026-01-31---v4-alpha)
