# 2025-11-18 - Beta V2
- Implement usage of _commercial_ `Aspose.Cells`
- Switch to `EPPlus-LPGL` (It's faster than v8.x ;-))
    - [Performance on 2025-11-28](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-28)


# 2025-11-27 - Beta V2
- Change benchmarks to use `ToString` to be fair on `ClosedXML`
- Attempt to use `FastExcel` - Failed due to Bug
- Use _commercial_ `FreeSpire`
    - [Performance on 2025-11-27](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-27)

# 2025-11-25  - Beta V2
- Make `DefinedName`'s work with `localSheetId`definitions
- Start to Add Benchmarks for range extraction
    - [Performance on 2025-11-25](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-25)
 
# 2025-11-19 - Beta V2
- Add `IEnumerable`s _All_ the way down ⤵️
    - i.e. remove the need for Asynchronous awaits
    - 🚀 Yielding More Performance improvements
        - [Performance on 2025-11-16](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-16)
    - ⛓️‍💥 **Breaking Change** 🔩
        - The Async classes now have `Async` appended to be distinct from the non async versions
        - But, `Async` inherit from the non, so they are interchangable

# 2025-11-16 - Beta V2
- Read `definedName`s (Ranges / Cell / Value / Dynamic) 📇
- Implement RangeExtraction
    - Global rangeNames
- Deal with blank rows in a sheet 🗋
    - Return a `null` cell row
- Deal with Empty cells in a row 🗅
    - Return a `null` cell
- Remove some warnings

# 2025-11-14 - V1
- 🚀 [Performance on 2025-11-12](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-12)
