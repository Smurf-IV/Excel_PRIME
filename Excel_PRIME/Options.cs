using System;

namespace ExcelPRIME;


/// <summary>
/// Specify how the internals will deal and expose the Cell type value
/// </summary>
public enum CellConversion
{
    /// <summary>
    /// Do not use
    /// </summary>
    [Obsolete("Do not use")]
    Unknown = 0,

    /// <summary>
    /// default  - Normally the Fastest option, Will leave the value as the read type (i.e. as string from XLSX).
    /// </summary>
    None,

    /// <summary>
    /// Will convert to the CLR type specified by the celltype.
    /// </summary>  
    ExcelCellType, 

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
