using System;

namespace ExcelPRIME;

public enum FileType
{
    [Obsolete("Do not use")]
    unknown = 0,

    Xlsx = 1,
    Xlsb
};

public record Options
{
    /// <summary>
    /// In the future this may be set, to allow the Open Xml cell type to be used in the return object
    /// </summary>
    internal bool ConvertToSpecifiedCellType { get; }
}
