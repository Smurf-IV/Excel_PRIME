# ExcelFAST 🌟
- **Excel F**oward **A**rraySegment **S**peedy **T**ransfer
- Without using any external libraries.

# What does that mean?
- _Yet another Excel reader ?_, 
    - Starting with .Net 8 as the performant Runtime (See Benchmarks)
    - .Net9 gives an extra 5% boost, 
    - .Net10 Another 5% over .Net9 ;-)
    - .Net11 Another 5% -> 10% (`async`'s ;-)) over .Net10 ;-)

Lets take each of the above elements and explain:

## Excel 📈
- Open _Large_ 2007 (Onwards) XLS**X** file formats and XLS**B**
- Zip Deflate format _Only_

## Foward
- Force the usage of minimal memory,
- Only 1 sheet open at a time
- Read only
- Only unzip the sheet(s) when they are asked for
- Only load the shared strings upto the current request number

## ArraySegment / ReadOnlySpan📦
- This is to be targetted at memory restricted environments (i.e. ASP Net VM's)
- So each of the returned rows are _not_ copied into a new array, but instead returned as an `ReadOnlySpan<Cell>` of the original array, so that the GC does not have to do any extra work.
- There are also enforcements to ensure that
    - The `ReadOnlySpan` is _not_ used outside of the current row, and that the next row will _invalidate_ the previous row's `ReadOnlySpan` (i.e. it will be overwritten with the next row's data)
    - Single `Cells` returned do not have scope outside of the caller (i.e. ref Structs), so the compiler can give warnings
    - With `ref struct`:
        - Compiler prevents storing in fields, arrays, or heap allocations
        - Caller cannot accidentally hold onto it past its valid lifetime
        - Type system enforces the "transient view" contract

## Speedy 🚀
- _Try_ **to be** as fast as possible, i.e.
    - Forward only Lazy loading
    - Only "Quick" decipher / convert of the cell(s) types to ease GC pressure
    - No attempt at "creating / using" datatables with headers etc.
    - Allow `CancellationToken`s to be used to allow page transitioning cancellation (More on this later)
- Strongly-typed accessors (`AsInt32`, `AsDateTime`, etc)
### `CancellationToken`s
- This is to allow the Large files to be _Aborted_
### IDisposable
- Got to tidy up those `Temp File`s, and release the `FileStream`'s
- Use the OS's `Temp File` caching, so if the memory is _tight_ then the Owner app will not have to worry about OOM exceptions, or having to use Swap Disk speeds.
### Q & A's
- Q: There are others that are faster
- A: Agreed, but then 
    - They do **not** have range extraction.
    - Or `optionally` allow the use of the OS's _TempFile System_ to store massive sheets

## Transfer 🌐
- The reason for this project, is to handle very large XSLX files (i.e. > 500K rows with > 180 columns per sheet, with multiple sheets of this size)
- For `ETL` validation scenarios, i.e. make sure that the user modified data that has been transferred has interaction rules applied, before moving onto the `T` and `L` stages
- Try not to hit / store in the LOH
- No internal .Net memory of previously loaded sheets / rows.
### Q & A's
- Q: 
- A: 
-----

# Caveats ⛔:
## Will _Only_ Open 1 sheet at a time 🤖
- Save on memory and threading overhead of external disposable(s)
## Will _Not_ Do: Poco 🤖
- A POCO / Type populator (Extensions can be written for that later)
## Will _Not_ Be a: Writer / Modifier 📚
- Totally beyond the scope of this project remit
-----