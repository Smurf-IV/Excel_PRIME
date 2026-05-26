# 2026-05-26 - V5 - Beta
- ✅ Cell object type 📅
    - ✅ Store cell _style_ type (see Options enum)
    - ✅ Unit Tests
    - ✅ Implement reading of the styles to determine the default `DateTime` / `DateOnly` / `TimeOnly` formats #19
- **Code-Level Optimizations**
   - ✅ Implement `ISpanFormattable` in CellValue
   - ✅ Optimisationm in the Xlsb workflow
   - ✅ Return to the usage of the FieldOffsets to store the BCL type to prevent boxings in the hot paths
   - ✅ Usage of the Fast convertors *i.e. ToDecimal is 3 times faster than Convert.ToDecimal* #20
- **Advanced Scenarios**
   - ✅ Enable `PublishTrimmed=true` with trim warnings resolved
   - ✅ Native AOT compilation testing
- **Bug Fixes**
    - Implement reading of the styles to determine the default `DateTime` / `DateOnly` / `TimeOnly` formats #19
    - `AsDecimal` method has an issue where it produces incorrect precision, but only with default options #20
    - When Attempting to use the "SkipRows" on a a sheet that has null rows to start with, causes infinite loop #27
    - When opening the source file, then use "Sharing Mode" to allow it to be opened by other things! (i.e. 2 instances of this !) #28

# 2026-05-12 - V4 - Bug fixes
- Return null when `EndValue` is spotted in the xml #22
- Return null for not found SheetId #23
- Changed CellValue to an abstract base class and removed all [FieldOffset(0)] fields.
- Introduced internal sealed class CellValue<T> : CellValue to store values of type T without boxing for value types.

# 2026-05-05 - V4
- ⛓️‍💥 **Breaking Change(s)**
    - Removal of `GetSheetFileName(int offsetSheetId);`
    - Removal of `GetDefinedRange` via `int sheetId`
    - Removal of `Index` property from `ISheet`
    - Internal Creation of WorkBooks
    - Internal implementation of `IOpenXmlWorkBookReader::GetSheetNames` now returns the relative path to the sheetName
    - `CellValue` is now a `class`, therefore no need to use `.Value`
    - `ICell.CellValue` is now nullable
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
- Implement `System.DBNull` return option, for empty cells
- etc.