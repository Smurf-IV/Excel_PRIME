using System;

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
    None,   // default  - Normally the Fastest option, Will leave the value in an IMemory<char> object (unless it is already a string)
    Number, // Sometimes the Faster option, But always smaller memory, Will attempt to convert to the nearest integral signed number type, i.e. int -> long -> decimal -> double
    [Obsolete("Not Implemented yet!")]
    NumberAndDates, // As number and will also "Have a go" at detecting dates type (DateTime, DateOnly, TimeOnly, TimeSpan)
    [Obsolete("Not Implemented yet!")]
    ForceStyles // As  Dates, and also takes into account the number of decimal places etc. from the style when converting / formatting
}

public record Options
{
    /// <summary>
    /// In the future this may be set, to allow the Open Xml cell type to be used in the return object
    /// </summary>
    public CellConversion CellConversionType { get; init; } = CellConversion.None;

    /// <summary>
    /// If you are only reading the sheets once, then _do_not_ use the OS TempFile
    /// </summary>
    /// <remarks>
    /// `false`: Useful if going to use this again, and sheets are big, or Multiple sheets are opening, and multithreaded
    /// `true`: Default, Use the internal rented buffer from the zipArchive
    /// </remarks>
    public bool UseSheetsOnlyOnce { get; init; } = true;
}
