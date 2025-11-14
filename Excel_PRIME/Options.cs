using System;

namespace ExcelPRIME;

/// <summary>
/// Specify the type of stream being opened / used
/// </summary>
public enum FileType
{
    [Obsolete("Do not use")]
    Unknown = 0,

    Xlsx = 1,

    [Obsolete("Do not use = Yet ;-)")]
    Xlsb
}

/// <summary>
/// Specify how the internals will deal and expose the Cell type value
/// </summary>
public enum CellConversion
{
    [Obsolete("Do not use")]
    Unknown = 0,

    /// <summary>
    /// default  - Normally the Fastest option, Will leave the value as a string.
    /// </summary>
    /// 
    None,
    /// <summary>
    /// Will attempt to convert to the nearest integral signed number type, i.e. int -> long -> decimal -> double. Dates will be converted from ISO 8601.
    /// </summary>
    Number,

    /// <summary>
    /// As number and will also "Have a go" at detecting the date type (DateTime, DateOnly, TimeOnly, TimeSpan)
    /// </summary>
    [Obsolete("Not Implemented yet!")]
    NumberAndDates,

    /// <summary>
    /// As NumberAndDates, and also takes into account the number of decimal places etc. from the style when converting / formatting
    /// </summary>
    [Obsolete("Not Implemented yet!")]
    ForceStyles 
}

/// <summary>
/// Specify how the internals deal with conversion and sheet access
/// </summary>
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
    /// `false`: Useful if going to use this again, and sheets are big, Or Multiple sheets are opening, Or multithreaded
    /// `true`: Default, Use the internal rented buffer from the zipArchive; Therefore single threaded access to Excel file.
    /// </remarks>
    public bool AccessExcelFileInForwardOnlyMode { get; init; } = true;
}
