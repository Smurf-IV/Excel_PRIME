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
    - [>] 👟 [2nd Pass Performance on 2025-12-21](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-21)


# 2025-12-10 - RC
- User defined, using the `"A1:B10"` or `"$A$1:$B$10"` syntax
  - [Range Performance on 2025-12-10](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-10)

# 2025-12-07 - Beta V2
- Investigate _memory usage_(s) 🧑‍💻
    - Removed finalizers from several classes
    - Added a lightweight Row pooling strategy
- More performance improvements 🏃‍➡️ [Performance on 2025-12-04](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-04)
- Replaced the `ReadString` implementation (Memory optimisation)
- Some code Cleanup
- Prefer returning ReadOnlyMemory<char> to avoid allocating a new string
- Sacrificed a little speed..
    - [Performance on 2025-12-07](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-07)

# 2025-12-01 - Beta V2
- Investigate _memory usage_(s) 🧑‍💻
    - Replaced the dictionary with a fixed-size Cell?[] 
    - Defensive bounds checks when assigning parsed cells to avoid out-of-range writes.
    - `Row` disposal when going out of `yield` scopes
    - Add `ThreadStringBuilderPool` for memory efficiency
    - Add `AccessPivotTable` and explain why the other libraries do **not work**
- [Performance on 2025-12-01](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-01)

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
