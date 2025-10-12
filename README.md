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

# Targets
## Phase 0
- [x] Setup this github
- [x] Create the main project
- [x] Add Unit Test project
- [x] Add simple Test Data

## Phase Alpha
- [x] Use Net Core Interface(s)
    - [x] Use `ZipArchive`
    - [x] Use `XDocument`
- [x] Implement Open / Dispose (Async)
    - [x] Sheet Names
    - [x] Shared Strings
- [x] Implement Sheet loading (unzip and be ready for use)
    - [x] Use `XDocument`
- [x] Implement Row extraction 
    - [x] Skip
    - [x] Delayed read - until a cell is actually needed
    - [x] Deal with Null / Empty cells (Utilise sparse array)
    - [x] Keep last used offset (i.e. no need to reload sheet if the next range API `startRow` call is later)

## Phase Beta
- [ ] Benchmarks
- [ ] Implement Sheet loading via `XmlReader.Create`
- [ ] More UnitTests

## Phase 1 - MVP
- [ ] More Benchmarks
- [ ] Read `definedName`s (Ranges)
    - [ ] Store from global
- [ ] Implement Sheet loading of
    - [ ] Q: Are there _Shared strings_ per sheet as well ?
    - [ ] Store from Local sheets (When opened)
- [ ] Implement Row extraction 
    - [ ] Allow ColumnHeader addressing (i.e. `ABF`)
- [ ] Implement RangeExtraction
    - [ ] Global rangeNames
    - [ ] Per Sheet rangeNames
    - [ ] User defined

## Phase 2 - Multi project deployments (Nuget)
- [ ] More Benchmarks
    - [ ] Add Other "Excel readers" to the Benchmark project(s)
- [ ] Cell Information
    - [ ] Value type indicator
    - [ ] Formatter applied
    - [ ] Anything else useful
- [ ] Investigate a different way of storing the _Shared strings_ to the Filesystem, when they are in the MB's
- [ ] Implement Interface for other Libs (Xml / Zip)

## Phase 3
- [ ] XSL**B**


## Phase 4 - Extension(s)
- [ ] Add option class to allow _Basic_ Cell value type identification
    - [ ] Extract into those types
    - [ ] Deal with `DateOnly` / `TimeOnly` fields
    - [ ] Use of user defined column schema (Excel Number Format nuget?)

- [ ] More ideas to be added later ;-)
