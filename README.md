# Excel_PRIME
**Excel** **P**erformant **R**eader via **I**nterfaces for **M**emory **E**fficiency.

# What does that mean?
Lets take each of the above elements and explain:

## Excel
- Open _Large_ 2007 (Onwards) XLSX file formats (Binary later maybe)

## Performant
- Try to be as fast as possible, i.e.
    - No Attempting to decipher / convert the cell(s) types (Its all text in lx)
    - No attempting to create datatables with headers etc.
    - Use `IEnumerable`s and offset starts (Row / Column)
    - Allow `CancellationToken`s to be used to allow page transitioning cancellation (More on this later)

## Reader
Read only, therefore no calculation / formula calls

## Interfaces 
- Will use the DotNet core functionality by default
- But, if your target deployment allows for the use of native performant binaries, then via the use of interfaces these will be pluggable
    - i.e. Using `ZStdzip` for getting the data streams out of the compressed Excel file faster.
    - A faster / slimmer implementation for xml stream reading

## Memory
- The reason for this is to handle very large XSLX files (i.e. > 500K rows with > 180 columns per sheets, with multiple sheets of this size)
- For ETL validation scenarios, i.e. make sure that the user modified data that has been transferred has interaction rules applied, before moving onto the T and L stages
- Try not to hit / store in the LOH
- No caching of previously loaded rows (i.e. _like_ a forward only reader)

## Efficiency
- As hinted by the above statements, this is to be targetted into memory restricted environments (i.e. ASP Net VM)
- Use the OS's `Temp File` caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
- Only unzip the sheet(s) when they are asked for

## Etc.
### `CancellationToken`s
- This is to allow the Large files to be _Aborted_
- Make "Most" of the API's Asynchronous `Task`s
### IDisposable
- Got to tidy up those `Temp File`s, and release the `File Stream`

<hr />

# It will not be:
## Thread safe
- Initially it will **Not** be thread safe, because the xml reader will be locked to the sheet in use.
## Cell object type
- Cell conversion when read (i.e. you will know the type that you want, and you can convert it.)
- This could later become an option if the `XmlConvert` classes are efficient (Or via the interface specs)
## Poco
- A POCO / Type populator (Extensions can be written for that later)
## Writer / Modifier
- Totally beyond the scope of this project remit
<hr />

# Targets
## Phase 0
- [x] setup the github
- [ ] Create the main project
- [ ] Add Unit Test project
- [ ] Add simple Test Data

## Phase 1
- [ ] Implement Open / Dispose (Async)
- [ ] Implement Reader Identification (i.e. versions of Excel file)
- [ ] Implement Sheet loading (unzip and be ready for use)
- [ ] Implement Row extraction
    - [ ] Whole
    - [ ] Subset (With offset)
    - [ ] Deal with Null / Empty cells
    - [ ] Efficient count (Test it !)

## Phase 2
- [ ] Implement Interface(s)
- [ ] Add Benchmark project
- [ ] Add option class to allow _Basic_ Cell type identification
- [ ] Read `definedName`s (Ranges)
    - [ ] Store from global
    - [ ] Store from Local sheets
- [ ] Extract rows via Range

## Phase 3
- [ ] Ideas to be added later ;-)
