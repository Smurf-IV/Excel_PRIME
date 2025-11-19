# Excel_PRIME 🌟
- **Excel**_**P**erformant **R**eader via **I**nterfaces for **M**emory **E**fficiency.
- Without using any external libraries.
- Optimised for Range extraction.

# What does that mean?
- _Yet another Excel reader ?_, 
    - Starting with .Net 8 as the performant Runtime (See Benchmarks)
    - V9 gives an extra 5% boost, 
    - V10 Another 5% ;-)

Lets take each of the above elements and explain:

## Excel 📈
- Open _Large_ 2007 (Onwards) XLS**X** file formats (Binary later, _maybe_)

## Performant 🚀
- _Try_ **to be** as fast as possible, i.e.
    - Forward only Lazy loading
    - Only "Quick" decipher / convert of the cell(s) types to ease GC pressure
    - No attempt at "creating / using" datatables with headers etc.
    - Use `IEnumerable`s with initial offset starts (Row / Column)
    - Allow `CancellationToken`s to be used to allow page transitioning cancellation (More on this later)
### Q & A's
- Q: There are others that are faster
- A: Agreed, but then 
    - They do not have range extraction.
    - Or `optionally` allow the use of the OS's _TempFile System_ to store massive sheets
    - Or re-use of already extracted (massive) sheets
    - Or allow multiple sheets to be read at the same time 
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
    - AND there is work in place to allow multiple sheets to be read at the same time (Unlike some to of the others that use global memory to represent a row)
    - And it appears that the current benchmarks do not extract unless a `ToString` and a check on the result is used (Otherwise the Jit removes the unassigned dead code)

## Efficiency 📦
- As hinted by the above statements, this is to be targetted at memory restricted environments (i.e. ASP Net VM's)
- Use the OS's `Temp File` caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
- Only unzip the sheet(s) when they are asked for
- Only load the shared strings upto the current request number
### Q & A's
- Q: 
- A:


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