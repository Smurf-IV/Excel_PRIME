# 2026-01-## - V4 - Alpha
- ⛓️‍💥 **Breaking Change(s)**
    - Removal of `GetSheetFileName(int offsetSheetId);`
    - Removal of `GetDefinedRange` via `int sheetId`
    - Removal of `Index` property from `ISheet`
    - Internal Creation of WorkBooks
    - Internal implementation of `IOpenXmlWorkBookReader::GetSheetNames` now returns the relative path to the sheetName


# 2026-01-16 - V3 
- Remove some `AggressiveOptimization` and allow `i-cache` to do its job
- Implement "Hot-Paths" for cell type access
- Reduce some memory allocations for ReadOnly CellCollections
  - [2026-01-16](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-16)

# 2026-01-11 - V3
- Add XLS**B** 💾 (BIFF12)
- ⛓️‍💥 **Breaking Change(s)**
    - `FileType` has been removed, and Open via the Public class type
    - `IXmlReaderHelpers` has become `IOpenXmlReaderHelpers`, with slightly different methods
    - `IXmlWorkBookReader` has become `IOpenXmlWorkBookReader`
    - `IXmlSheetReader` has become `IOpenXmlSheetReader`
    - Removal of the Conversion options `Number###`
    - Changed `GetAllCells` to return `IReadOnlyList<ICell?>?`
        - Watch out for those null rows !
- 🚀 Big Performance improvements [2026-01-11](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-11)
﻿
# 2025-12-14 - V2
- ⛓️‍💥 **Breaking Change** 🔩
    - The Async classes now have `Async` appended to be distinct from the non async versions
    - But, `Async` inherit from the non, so they are interchangable
- Implement [GetUserRange(...)](https://github.com/Smurf-IV/Excel_PRIME/issues/7)
  - [Range Performance on 2025-12-14](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-14)
- User defined, using the `"A1:B10"` or `"$A$1:$B$10"` syntax
  - [Range Performance on 2025-12-10](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-10)
- Improve _memory usage_(s) 🧑‍💻
- Make `DefinedName`'s work with `localSheetId`definitions
- Benchmarks for range extraction
- Add `IEnumerable`s _All_ the way down ⤵️
    - i.e. remove the need for Asynchronous awaits
    - 🚀 Yielding More Performance improvements
        - [Performance on 2025-11-16](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-16)
- Read `definedName`s (Ranges / Cell / Value / Dynamic) 📇
- Implement RangeExtraction
    - Global rangeNames
- Deal with blank rows in a sheet 🗋
    - Return a `null` cell row
- Deal with Empty cells in a row 🗅
    - Return a `null` cell
- Remove some warnings
