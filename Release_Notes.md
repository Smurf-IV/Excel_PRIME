# 2025-11-19 - Beta V2
- Add `IEnumerable`s _All_ the way down ⤵️
    - i.e. remove the need for Asynchronous awaits
    - 🚀 Yielding More Performance improvements
    - ⛓️‍💥 **Breaking Change** 🔩
        - The Async classes now have `Async` appended to be distinct from the non async versions
        - But, `Async` inherit from the non, so they are interchangable

# 2025-11-16 - Beta V2
- Read `definedName`s (Ranges / Cell / Value / Dynamic) 📇
- Implement RangeExtraction
    - Global rangeNames
    - Per Sheet referenced range names
    - User defined
- Deal with blank rows in a sheet 🗋
    - Return a `null` cell row
- Deal with Empty cells in a row 🗅
    - Return a `null` cell
- Remove some warnings

# 2025-11-14 - V1
- Parallel Sheet Access
- ApplicationIcon
- Remove some warnings
- Nuget deployment V1


# 2025-11-12
- Optimise for `CellConversion.None` 💪
    - [Performance on 2025-11-12](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-12)

# 2025-11-08
- Investigation into the smallest function 💪
    - [Performance on 2025-11-08](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-08)

# 2025-11-07
- Add option to use the Stream from `ZipEntry`
- Explain the Excel cell types
