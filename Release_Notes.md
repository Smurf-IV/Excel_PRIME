# 2026-01-04 - V3 XLSB - RC1
- ⛓️‍💥 **Breaking Change(s)
    - Change `GetAllCells` to return `IReadOnlyList<ICell?>?`
- Add `NonClosingStream` to  make it self documenting
- Remove `BufferedStream` usage from xlsx
- Add more XLSB Benchmark readers
- Performance improvements 🚀
    - [2026-01-04](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-04)
- Remove some compile warnings
- Make use of `[MethodImpl(MethodImplOptions.AggressiveOptimization)]`
 
# 2026-01-02 - V3 XLSB-Beta
- ⛓️‍💥 **Breaking Change(s)
    - Removal of the Conversion options `Number###`
- Read `definedName`s (Ranges / Cell / Value / Dynamic) 📇
- Implement "On Demand" conversion
    - Slightly slower, but less memory pressure `xslb`
    - [2026-01-02](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-02)

# 2025-12-21 - V3 XLSB-Beta
- ⛓️‍💥 **Breaking Change(s)
    - `FileType` has been removed, and Open via the Public class type
    - `IXmlReaderHelpers` has become `IOpenXmlReaderHelpers`, with slightly different methods
    - `IXmlWorkBookReader` has become `IOpenXmlWorkBookReader`
    - `IXmlSheetReader` has become `IOpenXmlSheetReader`
- ✅ Implement Sheet loading 
- ✅ Implement Row extraction 
    - ✅ Skip
    - ✅ Delayed read - until a cell is actually needed
    - ✅ Deal with Null / Empty cells
- ✅ Cell object type 📅
    - ✅ 👟 [2nd Pass Performance on 2025-12-21](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-21)

﻿# 2025-12-14 - V2
- Implement [GetUserRange(...)](https://github.com/Smurf-IV/Excel_PRIME/issues/7)
  - [Range Performance on 2025-12-14](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-14)

# 2025-12-10 - V2 RC
- User defined, using the `"A1:B10"` or `"$A$1:$B$10"` syntax
  - [Range Performance on 2025-12-10](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-10)

# 2025-1#-## - Beta V2
- Improve _memory usage_(s) 🧑‍💻
- ⛓️‍💥 **Breaking Change** 🔩
    - The Async classes now have `Async` appended to be distinct from the non async versions
    - But, `Async` inherit from the non, so they are interchangable
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
