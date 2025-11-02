using System;
using System.Threading;

namespace ExcelPRIME;

public enum FileType
{
    [Obsolete("Do not use")]
    Unknown = 0,

    Xlsx = 1,

    [Obsolete("Do not use = Yet ;-)")]
    Xlsb
}

public enum CellConversion
{
    [Obsolete("Do not use")] 
    Unknown = 0,
    None,   // default  - Fastest option, Will leave the value in an IMemory<char> object (unless it is already a string)
    Simple, // Will convert to double / bool / date dependent on defined ExcelCell type
    Number, // Will attempt to convert to the nearest integral signed number type, i.e. int -> long -> decimal -> double
    [Obsolete("Not Implemented yet!")]
    NumberAndDates, // AS number and will also "Have a go" at detecting dates type (DateTime, DateOnly, TimeOnly)
    [Obsolete("Not Implemented yet!")]
    FromStyles // As  Dates, and also take into account the number of decimal places etc from the style
}

public record Options
{
    /// <summary>
    /// In the future this may be set, to allow the Open Xml cell type to be used in the return object
    /// </summary>
    public CellConversion CellConversionType { get; init; } = CellConversion.None;
}
