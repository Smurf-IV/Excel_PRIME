# Excel_PRIME 🌟
- **Excel**_**P**erformant **R**eader via **I**nterfaces for **M**emory **E**fficiency.
- Without using any external libraries.
- Optimised for Range extraction.
- ![Excel.png](./Excel.png)

# What does that mean?
- _Yet another Excel reader ?_, 
    - Starting with .Net 8 as the performant Runtime (See Benchmarks)
    - .Net9 gives an extra 5% boost, 
    - .Net10 Another 5% over .Net9 ;-)

Lets take each of the above elements and explain:

## Excel 📈
- Open _Large_ 2007 (Onwards) XLS**X** file formats and XLS**B** (BIFF12) in V3.##
- Zip Deflate format _Only_

## Performant 🚀
- _Try_ **to be** as fast as possible, i.e.
    - Forward only Lazy loading
    - Only "Quick" decipher / convert of the cell(s) types to ease GC pressure
    - No attempt at "creating / using" datatables with headers etc.
    - Use `IEnumerable`s with initial offset starts (Row / Column)
    - Allow `CancellationToken`s to be used to allow page transitioning cancellation (More on this later)
- Now the fastest in Real world usage [2025-11-19 onwards](https://github.com/Smurf-IV/Excel_PRIME/discussions/2#discussioncomment-15013658)
### Q & A's
- Q: There are others that are faster
- A: Agreed, but then 
    - They do **not** have range extraction.
    - Or optionally allow the use of the OS's _TempFile System_ to store massive sheets
    - Or **re-use** of already extracted (massive) sheets
    - Or allow multiple sheets to be read at the **same** time 
        - because others use global memory to represent the current row
        - Or have a single access into the Zip Excel file


## Reader 📋
- Read only
    - Therefore no calculations or updates to formula calls

## Interfaces 🏗️
- Will use the DotNet core functionality by default
- But, if your target deployment allows for the use of native performant binaries, then via the use of interfaces these will be pluggable
    - i.e. Using `Zlib.Net` for getting the data streams out of the compressed Excel file faster. (Or `SharpZipLib` / `PowerPlayZipper`)
    - A faster / slimmer implementation for xml stream reading (i.e. TurboXml)
- Allow the implementation of different source files (i.e. XLS**B**)
### Q & A's
- Q: Why?
- A: As mentioned above, this is to allow a developer to replace with external nugets that might perform better XML speed etc.

## Memory 🌐
- The reason for this project, is to handle very large XSLX files (i.e. > 500K rows with > 180 columns per sheet, with multiple sheets of this size)
- For `ETL` validation scenarios, i.e. make sure that the user modified data that has been transferred has interaction rules applied, before moving onto the `T` and `L` stages
- Try not to hit / store in the LOH
- No internal .Net memory of previously loaded sheets / rows.
### Q & A's
- Q: It appears that this uses more memory than other implementations
- A: Currently yes, but it is being optimised for _Range Extraction_, 
    - AND for allowing multiple rows (With cell data) to be stored in memory at the same time, (i.e. via `ToList()` call);
    - AND to allow multiple sheets to be read at the same time (Unlike some to of the others that use a single global memory to represent a row)
    - And it appears that the current benchmarks do not extract unless a `ToString` and a check on the result is used (Otherwise the Jit removes the unassigned dead code)
    - And, the memory used will actually be used in the ETL pipeline anyway, so it's just being truthful

## Efficiency 📦
- As hinted by the above statements, this is to be targetted at memory restricted environments (i.e. ASP Net VM's)
- Use the OS's "Temp File" caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
- Only unzip the sheet(s) when they are asked for
- Only load the shared strings upto the current request number
### Q & A's
- Q: Sometimes the `Async` _await_ s add too much overhead
- A: true, that is why there are also the equivalent base interfaces that perform the same functionality without the need for the `async await` overheads.


## Etc. 🔧
### `CancellationToken`s
- This is to allow the Large files to be _Aborted_
- Make "Most" of the "Net Cores'" API's Asynchronous `Task`s
### IDisposable
- Got to tidy up those `Temp File`s, and release the `FileStream`'s

### Challenges:
- CellValue instances are returned to users
- They must be thread-safe (multiple readers possible)
- Each "Cell Type" / "Cell Instance" / " Row Instance"(string, numeric, boolean, datetime, error) have different lifecycle requirements

-----

# Caveats ⛔:
## _It will not_ be: Same sheet thread instance safe 📊
- It will **Not** be _same sheet Instance_ thread safe, because the xml reader will be locked (Forward only) to the sheet in use.
    - but you **_can_** Open the sheet more than once, and have different threads running over it,
    - And you **_can_** have Parallel threads access the Excel file
    - Just remember to set `Options{ AccessExcelFileInForwardOnlyMode = false}`
## _It will not_ do: Dynamic Ranges ⚠️
- i.e. Ones that contain formulas:
    - `<definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>`
## _It will not_ do: Poco 🤖
- A POCO / Type populator (Extensions can be written for that later)
## _It will not_ be: Writer / Modifier 📚
- Totally beyond the scope of this project remit

-----

| Badge 🔄 | Area   |
|--------------------------- |-------------|
| [![.NET](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml/badge.svg?branch=main)](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml) | Release build and tests |

<a href="https://info.flagcounter.com/dxXK"><img src="https://s01.flagcounter.com/map/dxXK/size_l/txt_000000/border_CCCCCC/pageviews_1/viewers_0/flags_0/" alt="Flag Counter" border="0"></a>
-----
<!-- TOC-->
- [Targets 🎯](#targets-)
  - [Phase 0](#phase-0)
  - [Phase Alpha](#phase-alpha)
  - [Phase Beta - Benchmarks ⏱️](#phase-beta---benchmarks-)
  - [Phase 1 - MVP 🔍](#phase-1---mvp-)
  - [Phase 2 - RC](#phase-2---rc)
    - [V2 Changes ➡️ 2025-12-14](#v2-changes--2025-12-14)
  - [Phase V3 - XLS**B** 💾 (BIFF12)](#phase-v3---xlsb--biff12)
    - [V3 Changes ➡️ 2026-01-16](#v3-changes--2026-01-16)
  - [Phase V4 - Specific Cell value type(s) #️⃣](#phase-v4---specific-cell-value-types-)
  - [V4 changes ➡️ 2026-05-##](#v4-changes--2026-05-)
  - [Phase 5 - User Cell Value type formatting 💽 & Performance Optimizations 🏃‍➡️](#phase-5---user-cell-value-type-formatting---performance-optimizations-)
  - [Phase 6 - Third Party Nugets 📦](#phase-6---third-party-nugets-)
  - [Phase 7 - ideas 💡](#phase-7---ideas-)
<!-- TOC -->
-----

# Targets 🎯
## Phase 0
- ✅ Setup this github
- ✅ Create the main project
- ✅ Add Unit Test project
- ✅ Add simple Test Data

## Phase Alpha
- ✅ Use Net Core Interface(s)
    - ✅ Use `ZipArchive`
    - ✅ Use `XDocument`
- ✅ Implement Open / Dispose (Async)
    - ✅ Sheet Names
    - ✅ Shared Strings
- ✅ Implement Sheet loading (unzip and be ready for use)
    - ✅ Use `XDocument` as POC only
- ✅ Implement Row extraction 
    - ✅ Skip
    - ✅ Delayed read - until a cell is actually needed
    - ✅ Deal with Null / Empty cells (Utilise sparse array?)
    - ✅ Keep last used offset (i.e. no need to reload sheet if the next range API `startRow` call is later)

## Phase Beta - Benchmarks ⏱️
- ✅ Benchmarks
    - ✅ Add Other "Excel readers" to the Benchmark project(s)
    - 🎉 Now With `Sylvan.Data.Excel`
    - 🎉 Now With `XlsxHelper`
- ✅ More UnitTests

-----

## Phase 1 - MVP 🔍
- ✅ Add `IEnumerable`s and benchmark
- ✅ Implement `XmlReader.Create` for
- ✅ More Benchmarks
    - Now With `FastExcel`
    - ✅ Some Profiling Enahancements 
- ✅ Better Storage of the SharedStrings
- ✅ Cell object type 📅
- ✅ Use internal `ZipEntry` rented buffer
- ✅ Investigation into the smallest function 💪
- ✅ Optimise for `CellConversion.None` 💪
- ✅ Parallel Sheet threads Access
- ✅ Nuget
    - ✅ Beta etc.
    - 🎊 Released as Nuget V1.yyMM.dd -> **`1.2511.14`**

-----

## Phase 2 - RC
- ✅ Add `IEnumerable`s _All_ the way down ⤵️
- ✅ Nuget
    - ✅ Manual workflow deploy Release
    - ✅ Manual workflow deploy Beta
- ✅ Read "definedName"s (Ranges / Cell / Value / Dynamic) 📇
- ✅ Deal with blank rows in a sheet 🗋
- ✅ Deal with Empty cells in a row 🗅
- ✅ Implement Sheet scoping of "definedName"s
- ✅ Implement Row extraction 📟
- ✅ Implement RangeExtraction 📲
- ✅ Add Benchmarks for "Excel readers" That perform Range Extraction
    - ✅ `ClosedXML` Version="0.105.0"
    - ✅ `EPPlus_LPGL` Version="4.5.3.13"
    - ⚠️ `FastExcel` Version="3.0.13" -> [Fails on Range Extraction](https://github.com/ahmedwalid05/FastExcel/issues/89)
    - ✅ `FreeSpire.XLS` Version="14.2.0"
    - ✅ `Aspose.Cells` Version="25.11.0"
    - ⚠️ Extend benchmarks to cover the other large file types
        - It appears that most of the others do not like the `pivot-tables` file.!! 🤯
        - [Performance on 2025-11-28](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-28)
- ✅ Investigate _memory usage_(s) 🧑‍💻
    - ✅ Sacrificed a little speed ➡️ [Performance on 2025-12-07](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-07)
- ✅ Release as Nuget V2.2512-10 💨

-----

### V2 Changes ➡️ 2025-12-14
- Implement [GetUserRange(...)](https://github.com/Smurf-IV/Excel_PRIME/issues/7)
  - [Range Performance on 2025-12-14](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-14)

-----

## Phase V3 - XLS**B** 💾 (BIFF12)
- ⛓️‍💥 **Breaking Change(s)**
    - `FileType` has been removed, and Open via the Public class type
    - `IXmlReaderHelpers` has become `IOpenXmlReaderHelpers`, with slightly different methods
    - `IXmlWorkBookReader` has become `IOpenXmlWorkBookReader`
    - `IXmlSheetReader` has become `IOpenXmlSheetReader`
    - Removal of the Conversion options `Number###`
    - Changed `GetAllCells` to return `IReadOnlyList<ICell?>?`
        - Watch out for those null rows !
- ✅ Branch and beta yml
    - ✅ Convert test data in xls**b** format
- ✅ Implement Open / Dispose (Async)
    - ✅ Sheet Names
    - ✅ Shared Strings
- ✅ Implement Sheet loading 
- ✅ Implement Row extraction 
    - ✅ Skip
    - ✅ Delayed read - until a cell is actually needed
    - ✅ Deal with Null / Empty cells
- ✅ Cell object type 📅
- ✅ Benchmarks 🖲️
    - ✅ Add "Excel readers" That support XLS**B** Extraction
    - ✅ 🚶‍➡️ [1st Pass Performance on 2025-12-20](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-20)
    - ✅ 👟 [2nd Pass Performance on 2025-12-21](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-21)
- ✅ Read "definedName"s (Ranges / Cell / Value / Dynamic) 📇
    - ✅ Read from global
- ✅ Strongly-typed accessors (`AsInt32`, `AsDateTime`, etc)
    - Slightly slower, but less memory pressure for `xslb`
    - ✅ [2026-01-02](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-02)
- ✅ Parallel Sheet threads Access
    - ✅ Multiple times (with locking)
- ✅ Release as Nuget V3.yyMM.dd
    - 🎊 Released **RC1** as Nuget `V3.2601.04-RC1`
- ✅ Investigate Performance and edge cases, then Release as Stable
   - 🚀 Big Performance improvements [2026-01-11](Performance.md#2026-01-11)
- 🎊 Released **V3** as Nuget `V3.2601.11`

-----

### V3 Changes ➡️ 2026-01-16
- Remove some `AggressiveOptimization` and allow `i-cache` to do its job
- Implement "Hot-Paths" for cell type access
- Reduce some memory allocations for ReadOnly CellCollections
  - [2026-01-16](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-16)

-----

## Phase V4 - Specific Cell value type(s) #️⃣
- ⛓️‍💥 **Breaking Change(s)**
    - Removal of `GetSheetFileName(int offsetSheetId);`
    - Removal of `GetDefinedRange` via `int sheetId`
    - Removal of `Index` property from `ISheet`
    - Internal Creation of WorkBooks
    - Internal implementation of `IOpenXmlWorkBookReader::GetSheetNames` now returns the relative path to the "Sheet Name"
    - `CellValue` is now a `class`, therefore no need to use `.Value`
    - `ICell.CellValue` is now nullable
- ✅ Cell object type 📅
  - ✅ "Best Effort" `Operator` based conversion
  - ✅ TryGet`Type` will return `out type`, if stored as that type.
  - ✅ Unit Tests
- ✅ Performance
  - ✅ Use `ValueTask` and reduce memory allocations in some hot paths
  - 🚀 [Fix fallout from making `CellValue` is now a `class`](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md##2026-01-31-v4-alpha)
  - ✅ `ArrayPool` support has been added to ThreadStringBuilderPool using ArrayPool<char>.
  - ✅ Release-specific optimizations added
    - ✅ EnableTrimAnalyzer: true
    - ✅ TieredCompilation: true
    - ✅ TieredCompilationQuickJit: true
    - ✅ TieredCompilationQuickJitForLoops: true
- ✅ Implement `System.DBNull` return option, for empty cells
  - ✅ Implement `INullRow` return option, for empty rows
  - ✅ Update tests to use `INullRow` detection
- ✅ Implement `GetCell###(string columnLetters, ...)` #8
- 🚀 [2026-05-05](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-05-05---v4)

-----
## V4 changes ➡️ 2026-05-##
- Fix "Always get as date" issue for some cells #19


-----

## Phase 5 - User Cell Value type formatting 💽 & Performance Optimizations 🏃‍➡️
- ⛓️‍💥 **Breaking Change(s)**
    - None yet.
- [ ] Cell object type 📅
    - [ ] Store cell _style_ type (see Options enum)
    - [ ] Use of _user_ defined column schema
    - [ ] Formatter applied -> `CellConversion.ForceStyles`
    - [ ] Unit Tests
    - [ ] Deal with `DateOnly` / `TimeOnly` fields -> `CellConversion.NumberAndDates` 💹
**Code-Level Optimizations**
   - [ ] Implement `ISpanFormattable` in CellValue
   - [ ] Use `CollectionsMarshal` for zero-copy operations
   - [ ] Add `System.Runtime.Intrinsics` for SIMD
**Advanced Scenarios**
   - [ ] Enable `PublishTrimmed=true` with trim warnings resolved
   - [ ] Native AOT compilation testing
   - [ ] IAsyncEnumerable stream processing
**Monitoring**
   - [ ] Add performance regression tests
   - [ ] Implement ETW profiling in CI/CD
   - [ ] Track JIT compilation metrics

-----

## Phase 6 - Third Party Nugets 📦
- ⛓️‍💥 **Breaking Change(s)**
    - None yet.
- [ ] Excercise the Implementation of Interfaces for other Libs (Xml / Zip)
    - [ ] Separate Nuget(s) ?
- [ ] Benchmarks
  - [ ] e.g. search isages of `Class PoolingArrayBufferWriter<T>`
  - [ ] 

-----

## Phase 7 - ideas 💡
- [ ] Investigate a different way of storing the _Shared strings_ to the Filesystem, when they are in the MB's
  - [ ] e.g. Search for `Class FileBufferingWriter`
- [ ] Investigate possibility of using "Pipelining" to get data for Next row / cell population after yield?
    - [ ] Locking
    - [ ] How to deal with rows that are completely blank
    - [ ] `fibres` ?
- [ ] Indicate that things may be `Hidden` 🖺
    - [ ] Sheet
    - [ ] Row
    - [ ] Column
    - [ ] Cell ?
- [ ] Indicate that things may be `Readonly`
    - [ ] Sheet
    - [ ] Row
    - [ ] Column
    - [ ] Cell ?

- [ ] More ideas to be added later, Please suggest... ;-)
