# Excel_PRIME 🌟
- **Excel**_**P**erformant **R**eader via **I**nterfaces for **M**emory **E**fficiency.
- Without using any external libraries.
- Optimised for Range extraction.
- ![Excel.png](./Excel.png)

# What does that mean?
- _Yet another Excel reader ?_, 
    - Starting with .Net 8 as the performant Runtime (See Benchmarks)
    - V9 gives an extra 5% boost, 
    - V10 Another 5% ;-)

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
    - Or `optionally` allow the use of the OS's _TempFile System_ to store massive sheets
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
- A: Currently yes, but it is being optimised for `Range Extraction`, 
    - AND for allowing multiple rows (With cell data) to be stored in memory at the same time, (i.e. via `ToList()` call);
    - AND to allow multiple sheets to be read at the same time (Unlike some to of the others that use "a single" global memory to represent a row)
    - And it appears that the current benchmarks do not extract unless a `ToString` and a check on the result is used (Otherwise the Jit removes the unassigned dead code)

## Efficiency 📦
- As hinted by the above statements, this is to be targetted at memory restricted environments (i.e. ASP Net VM's)
- Use the OS's `Temp File` caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
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

-----

# It will **_not_** ⛔:
## Be: Same sheet Thread safe 📊
- It will **Not** be _same sheet Instance_ thread safe, because the xml reader will be locked (Forward only) to the sheet in use.
    - but you **_can_** Open the sheet more than once, and have different threads running over it,
    - And you **_can_** have Parallel threads access the Excel file
    - Just remember to set `Options{ AccessExcelFileInForwardOnlyMode = false}`
## Do: Dynamic Ranges ⚠️
- i.e. Ones that contain formulas:
    - `<definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>`
## Do: Poco 🤖
- A POCO / Type populator (Extensions can be written for that later)
## Be a: Writer / Modifier 📚
- Totally beyond the scope of this project remit

-----

| Badge 🔄 | Area   |
|--------------------------- |-------------|
| [![.NET](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml/badge.svg?branch=main)](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml) | Release build and tests |

<a href="https://info.flagcounter.com/dxXK"><img src="https://s01.flagcounter.com/map/dxXK/size_l/txt_000000/border_CCCCCC/pageviews_1/viewers_0/flags_0/" alt="Flag Counter" border="0"></a>
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
    - ⚠️ Still not convinced whether to implement "all the way down"
- ✅ Implement `XmlReader.Create` for
    - ✅ Loading sharedStrings
    - ✅ Sheet loading
    - ✅ Some Profiling Enahancements 
        - ✅ Performance [2025-10-18-pm](Performance.md#2025-10-18-pm)
- ✅ More Benchmarks
    - Now With `FastExcel`
    - ✅ Some Profiling Enahancements 
        - 🚀 Big Performance improvements [2025-10-19-pm](Performance.md#2025-10-19-pm)
- ✅ Better `Storage` of the SharedStrings
    - ✅ Use of LazyLoading Class
        - ⚠️ Performance [2025-10-14](Performance.md#2025-10-14)
    - ✅ Use of Derived `XmlNamedTable` implementations
    - ✅ Locking for separate sheet thread reading
        - ⚠️ Performance [2025-10-25](Performance.md##2025-10-25)
    - ✅ Restricted storage (i.e. do not return things that are not relevant)
      - 🚀 Big Performance improvements [2025-10-26](Performance.md#2025-10-26)
- ✅ Cell object type 📅
      - 🚀 Big Performance improvements [2025-11-01](Performance.md#2025-11-01)
    - ✅ Cell converted when read (i.e. you will know the type that you want, and you can convert it.)
      - 🚀 Big Performance improvements [2025-11-04](Performance.md#2025-11-04)
- ✅ Use internal `ZipEntry` rented buffer
    - ✅ Add and explain usage in options
      - 🚀 Big Performance improvements  [2025-11-07](Performance.md#2025-11-07)
- ✅ Investigation into the smallest function 💪
    - 🚀 More Performance improvements  [2025-11-08](Performance.md#2025-11-08)
- ✅ Optimise for `CellConversion.None` 💪
    - 🚀 More Performance improvements  [2025-11-12](Performance.md#2025-11-12)
- ✅ Parallel Sheet threads Access
    - ✅ Multiple times (with locking)
- ✅ Nuget
    - ✅ Beta etc.
    - 🎊 Released as Nuget V1.yyMM.dd -> **`1.2511.14`**
-----

## Phase 2 - RC
- ✅ Add `IEnumerable`s _All_ the way down ⤵️
    - i.e. remove the need for Asynchronous awaits
    - 🚀 Yielding More Performance improvements  [2025-11-19](Performance.md#2025-11-19)
    - ⛓️‍💥 **Breaking Change** 🔩
        - The Async classes now have `Async` appended to be distinct from the non async versions
        - But, `Async` inherit from the non, so they are interchangable
- ✅ Nuget
    - ✅ Manual workflow deploy Release
    - ✅ Manual workflow deploy Beta
- ✅ Read `definedName`s (Ranges / Cell / Value / Dynamic) 📇
    - ✅ Read from global
    - ✅ Handle Dynamics (i.e. do not fall over! 🤷)
- ✅ Deal with blank rows in a sheet 🗋
    - ✅ Return a `null` cell row
- ✅ Deal with Empty cells in a row 🗅
    - ✅ Return a `null` cell (e.g. `<c r="F12" s="8"/>`)
- ✅ Implement Sheet scoping of `definedName`s
    - ✅ i.e. `<definedName name="OrderSize" localSheetId="0">'Try it Yourself'!$C$12:$E$12</definedName>`
    - Note: The above will be referenced as `OrderSize (Try it Yourself)` as shown in LibreOffice.
- ✅ Implement Row extraction 📟
    - ✅ Allow ColumnHeader addressing (i.e. start -> end columns)
- ✅ Implement RangeExtraction 📲
    - ✅ Global rangeNames
    - ✅ Make `DefinedName`'s work with `localSheetId`definitions
    - ✅ User defined, using the `"A1:B10"` or `"$A$1:$B$10"` syntax
      - [Range Performance on 2025-12-10](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-10)
- ✅ Add Benchmarks for "Excel readers" That perform Range Extraction
    - ✅ `ClosedXML` Version="0.105.0"
        - [Performance on 2025-11-25](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-25)
    - ✅ `EPPlus_LPGL` Version="4.5.3.13"
        - [Performance on 2025-11-25](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-25)
    - ⚠️ `FastExcel` Version="3.0.13" -> [Fails on Range Extraction](https://github.com/ahmedwalid05/FastExcel/issues/89)
    - ✅ `FreeSpire.XLS` Version="14.2.0"
        - [Performance on 2025-11-27](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-27)
    - ✅ `Aspose.Cells` Version="25.11.0"
        - [Performance on 2025-11-28](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-11-28)
    - ⚠️ Extend bencmarks to cover the other large file types
        - It appears that most of the others do not like the `pivot-tables` file.!! 🤯
- ✅ Investigate _memory usage_(s) 🧑‍💻
    - ✅ Some performance improvements 🏃‍➡️ [Performance on 2025-12-01](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-01)
    - ✅ More performance improvements 🏃‍➡️ [Performance on 2025-12-04](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-04)
    - ✅ Sacrificed a little speed ➡️ [Performance on 2025-12-07](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2025-12-07)
- ✅ Release as Nuget V2.2512-10 💨
-----

## Phase 3 - XLS**B** 💾 (BIFF12) - Beta V3
- ⛓️‍💥 **Breaking Change(s)**
    - `FileType` has been removed, and Open via the Public class type
    - `IXmlReaderHelpers` has become `IOpenXmlReaderHelpers`, with slightly different methods
    - `IXmlWorkBookReader` has become `IOpenXmlWorkBookReader`
    - `IXmlSheetReader` has become `IOpenXmlSheetReader`
    - Removal of the Conversion options `Number###`
    - Changed `GetAllCells` to return `IReadOnlyList<ICell?>?`
        - Watch out for those null rows !
- ✅ Branch and beta yml
    - ✅ Convert test data in xls**b** format (External `ultra - deflate` Recompress)
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
- ✅ Read `definedName`s (Ranges / Cell / Value / Dynamic) 📇
    - ✅ Read from global
- ✅ Strongly-typed accessors (`AsInt32`, `AsDateTime`, etc)
    - Slightly slower, but less memory pressure for `xslb`
    - ✅ [2026-01-02](https://github.com/Smurf-IV/Excel_PRIME/blob/main/Performance.md#2026-01-02)
- ✅ Parallel Sheet threads Access
    - ✅ Multiple times (with locking)
- [>] Release as Nuget V3.yyMM.dd
    - 🎊 Released **RC1** as Nuget `V3.2601.04-RC1`
    - [ ] Investigate Performance and edge cases, then Release as Stable
-----

## Phase 4 - Specific Cell value type(s) #️⃣
- [ ] Cell object type 📅
    - [ ] `Operator` based conversion
    - [ ] Deal with `DateOnly` / `TimeOnly` fields -> `CellConversion.NumberAndDates` 💹
    - [ ] Use of user defined column schema (Excel Number Format nuget?)
    - [ ] Formatter applied -> `CellConversion.ForceStyles`
    - [ ] Investigate if the `XmlConvert` classes are efficient
- [ ] Benchmarks
-----

## Phase 5 - Third Party Nugets 📦
- [ ] Excercise the Implementation of Interfaces for other Libs (Xml / Zip)
    - [ ] Separate Nuget(s) ?
- [ ] Benchmarks
-----

## Phase 6 - ideas 💡
- [ ] Investigate a different way of storing the _Shared strings_ to the Filesystem, when they are in the MB's
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
