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

# 2025-11-08
- Investigation into the smallest function 💪

# 2025-11-07
- Add option to use the Stream from `ZipEntry`
- Explain the Excel cell types
- Make the Benchmark "File Names" clearer

# 2025-11-04
- Move Rented buffer into Row
- Use `char[]` for ColumnName storage
- Tinker with `DefinedRange` class
- Add styles of extraction to benchmark

# 2025-11-01
- Make all the benchmarks perform a `ToString` to ensure data is actually retrieved
- Pass `Atomized` strings around
- Pass `InstanceContext` around
- Add `None` as the default convert option
- Remove File buffers to allow OS to perform caching
- Correct Sheet Id's
- Extend options for Value extraction

# 2025-11-01
- Make all the benchmarks perform a `ToString` to ensure data is actually retrieved
- Pass `Atomized` strings around
- Pass `InstanceContext` around
- Add `None` as the default convert option
- Remove File buffers to allow OS to perform caching
- Correct Sheet Id's
- Extend options for Value extraction

# 2025-10-25
- make use of NameTable and `Object.ReferenceEquals`
- Allow many threads to add to the SharedStrings
- Add basic rawValue type detection
- Lowered the memory footprint

# 2025-10-26
- Investigate the usage of a derived `XmlNameTable`
- Some tweaks to lifetimes

# 2025-10-23
- After the fixes in the LazyStringLoader (Now appear to have more memory used??)

# 2025-10-19 pm
- Removed the "ToString()" call in the Benchmark, because the others did not have it!
- More profilling to see where things can be saved
- Now With FastExcel (Will not be using it again !)

# 2025-10-18 pm
- All Cell Access
- Some performance Analysis and changes

# 2025-10-18 am
- Loading test
- I'll not bother with these anymore, as they do not reflect "Real" usage

# 2025-10-17
- Shows that the strings are now lazy loaded
- Still not as good as `XlsxHelper` for the just file loading

# 2025-10-14
- This set of tests accessed every cell of the returned rows
- Therefore excercises the retrieval of all the `SharedStrings`

# 2025-10-13
- This set of tests "Only" accessed the first cell of the returned rows
- Therefore did not really excercise the retrieval of the `SharedStrings`, etc.

# 2025-10-08
- Can the large files be loaded
