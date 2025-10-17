# Excel_PRIME
**Excel**_**P**erformant **R**eader via **I**nterfaces for **M**emory **E**fficiency.

# What does that mean?
_Yet another Excel reader ?_, but starting with .Net 8 as the performant Runtime.

Lets take each of the above elements and explain:

## Excel
- Open _Large_ 2007 (Onwards) XLSX file formats (Binary later maybe)

## Performant
- Try to be as fast as possible, i.e.
    - Forward only Lazy loading
    - No Attempting to decipher / convert the cell(s) types (Its all text in lx)
    - No attempting to create /use datatables with headers etc.
    - Use `IEnumerable`s with initial offset starts (Row / Column)
    - Allow `CancellationToken`s to be used to allow page transitioning cancellation (More on this later)

## Reader
Read only, therefore no calculation / formula calls

## Interfaces 
- Will use the DotNet core functionality by default
- But, if your target deployment allows for the use of native performant binaries, then via the use of interfaces these will be pluggable
    - i.e. Using `Zlib.Net` for getting the data streams out of the compressed Excel file faster.
    - A faster / slimmer implementation for xml stream reading (i.e. TurboXml)

## Memory
- The reason for this is to handle very large XSLX files (i.e. > 500K rows with > 180 columns per sheet, with multiple sheets of this size)
- For `ETL` validation scenarios, i.e. make sure that the user modified data that has been transferred has interaction rules applied, before moving onto the `T` and `L` stages
- Try not to hit / store in the LOH
- No internal caching of previously loaded sheets / rows.

## Efficiency
- As hinted by the above statements, this is to be targetted at memory restricted environments (i.e. ASP Net VM's)
- Use the OS's `Temp File` caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
- Only unzip the sheet(s) when they are asked for

## Etc.
### `CancellationToken`s
- This is to allow the Large files to be _Aborted_
- Make "Most" of the API's Asynchronous `Task`s
### IDisposable
- Got to tidy up those `Temp File`s, and release the `File Stream`

<hr />

# It will **_not_** be:
## Thread safe
- Initially it will **Not** be _same sheet_ thread safe, because the xml reader will be locked to the sheet in use.
## Cell object type
- Cell converted when read (i.e. you will know the type that you want, and you can convert it.)
- This could later become an option if the `XmlConvert` classes are efficient (Or via the interface specs)
## Poco
- A POCO / Type populator (Extensions can be written for that later)
## Writer / Modifier
- Totally beyond the scope of this project remit

<hr />

| Badge   | Area   |
|--------------------------- |-------------|
| [![.NET](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml/badge.svg?branch=main)](https://github.com/Smurf-IV/Excel_PRIME/actions/workflows/dotnet.yml) | Release build and tests |

<hr />

# Targets
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
    - ✅ Use `XDocument`
- ✅ Implement Row extraction 
    - ✅ Skip
    - ✅ Delayed read - until a cell is actually needed
    - ✅ Deal with Null / Empty cells (Utilise sparse array)
    - ✅ Keep last used offset (i.e. no need to reload sheet if the next range API `startRow` call is later)

## Phase Beta - Benchmarks
- ✅ Benchmarks
    - ✅ Add Other "Excel readers" to the Benchmark project(s)
- ✅ More UnitTests
    - ⚠️ Performance [2025-10-08](Performance.md#2025-10-08)

<hr />

## Phase 1 - MVP
- [>] Add Non `IAsyncEnumerable`s and benchmark
    - ⚠️ Performance [2025-10-13](Performance.md#2025-10-13)
- [>] Implement `XmlReader.Create` for
    - [x] Loading sharedStrings
        - ⚠️ Performance [2025-10-14](Performance.md#2025-10-14)
    - [»] Sheet loading
- ✅ Better `Storage` of the SharedStrings
    - ✅ Use of LazyLoading Class
        - ⚠️ Performance [2025-10-14](Performance.md#2025-10-14)
- [ ] More Benchmarks
- [ ] Read `definedName`s (Ranges)
    - [ ] Store from global
- [>] Implement Sheet loading of
    - ✅ Multiple times with locking
    - [ ] Store `definedName` from Local sheets (When opened)
- [ ] Implement Row extraction 
    - [ ] Allow ColumnHeader addressing (i.e. `ABF`)
- [ ] Implement RangeExtraction
    - [ ] Global rangeNames
    - [ ] Per Sheet rangeNames
    - [ ] User defined

## Phase 2 - Multi project deployments (Nuget)
- [ ] More Benchmarks
    - [ ] Add even more "Excel readers" to the Benchmark project(s)
- [ ] Indicate that things may be `Hidden`
    - [ ] Sheet
    - [ ] Row
    - [ ] Column
    - [ ] Cell ?
- [ ] Indicate that things may be `Readonly`
    - [ ] Sheet
    - [ ] Row
    - [ ] Column
    - [ ] Cell ?
- [ ] Cell Information
    - [ ] Value type indicator
    - [ ] Formatter applied
    - [ ] Anything else useful
- [ ] Investigate a different way of storing the _Shared strings_ to the Filesystem, when they are in the MB's
- [ ] Implement Interface for other Libs (Xml / Zip)

## Phase 3
- [ ] XLS**B**


## Phase 4 - Extension(s)
- [ ] Add option class to allow _Basic_ Cell value type identification
    - [ ] Extract into those types
    - [ ] Deal with `DateOnly` / `TimeOnly` fields
    - [ ] Use of user defined column schema (Excel Number Format nuget?)
- [ ] Investigate possibility of using "Pipelining" to get data for Next row / cell population after yield?
    - [ ] Locking
    - [ ] How to deal with rows that are completely blank
    - [ ] `fibres` ?

- [ ] More ideas to be added later, Please suggest... ;-)
