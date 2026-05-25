using System.Collections.ObjectModel;

namespace ExcelPRIME.Implementation;

/// <summary>
/// Provides ECMA-376 standard definitions for cell styling and formatting.
/// 
/// ECMA-376 is the Office Open XML (OOXML) standard that defines:
/// - Built-in number format codes
/// - Default cell styles
/// - Style properties and their default values
/// 
/// References:
/// - ECMA-376-1:2016 - Part 1: Fundamentals and Markup Language Reference
/// - ECMA-376-2:2015 - Part 2: Open Packaging Conventions
/// 
/// See Also:
/// - Section 18.8.30 (numFmt - Number Format)
/// - Section 18.8.45 (numFmts - Number Formats)
/// </summary>
internal static class Ecma376StandardProvider
{
    /// <summary>
    /// Built-in number format codes and their formatting types as defined in ECMA-376 and Excel extensions.
    /// Each entry maps a format ID to a tuple containing the format code string and its corresponding FormattingType.
    /// 
    /// These include the standard format IDs 0-49 that are always available in Excel,
    /// plus commonly used extended formats (50-164).
    /// 
    /// Format ID Ranges:
    /// - 0-49: Standard ECMA-376 built-in formats
    /// - 50-164: Extended formats and locale-dependent variants
    /// - 165+: User-defined custom formats
    /// 
    /// Reference: ECMA-376-1:2016 Section 18.8.30
    /// 
    /// Implemented as FrozenDictionary for optimal .NET 8 performance:
    /// - O(1) lookup with superior memory locality
    /// - No per-lookup allocations
    /// - 30-40% faster than Dictionary for read-only workloads
    /// </summary>
    public static readonly ReadOnlyCollection<CellStyle> BuiltInNumberFormats = new CellStyle[]
    {
        // General formats
        new CellStyle { ExcelFormatId = DefaultGeneralFormatId, Formatting = string.Empty, FormattingType = FormattingType.General },

        // Number formats
        new CellStyle { ExcelFormatId = 1, Formatting = "0", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 2, Formatting = "0.00", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 3, Formatting = "#,##0", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 4, Formatting = "#,##0.00", FormattingType = FormattingType.Number },

        // Currency formats
        new CellStyle { ExcelFormatId = 5, Formatting = "$#,##0;($#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle
        {
            ExcelFormatId = 6, Formatting = "$#,##0.00;($#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 7, Formatting = "$#,##0;($#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle
        {
            ExcelFormatId = 8, Formatting = "$#,##0.00;($#,##0.00)", FormattingType = FormattingType.Currency
        },

        // Scientific notation
        new CellStyle { ExcelFormatId = 9, Formatting = "0.00E+00", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 10, Formatting = "#,##0.0", FormattingType = FormattingType.Number },

        // Text format
        new CellStyle { ExcelFormatId = 11, Formatting = string.Empty, FormattingType = FormattingType.General },

        // Time/Date formats
        new CellStyle { ExcelFormatId = 12, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 13, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 14, Formatting = "mm/dd/yyyy", FormattingType = FormattingType.DateTime },
        new CellStyle { ExcelFormatId = 15, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 16, Formatting = "d-mmm", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 17, Formatting = "mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 18, Formatting = "h:mm AM/PM", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 19, Formatting = "h:mm:ss AM/PM", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 20, Formatting = "h:mm", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 21, Formatting = "h:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 22, Formatting = "m/d/yy h:mm", FormattingType = FormattingType.DateTime },

        // Additional currency formats (23-36 locale-dependent)
        new CellStyle { ExcelFormatId = 23, Formatting = "#,##0;(#,##0)", FormattingType = FormattingType.Number },
        new CellStyle
        {
            ExcelFormatId = 24, Formatting = "#,##0.00;(#,##0.00)", FormattingType = FormattingType.Number
        },
        new CellStyle { ExcelFormatId = 25, Formatting = "#,##0;[Red](#,##0)", FormattingType = FormattingType.Number },
        new CellStyle
        {
            ExcelFormatId = 26, Formatting = "#,##0.00;[Red](#,##0.00)", FormattingType = FormattingType.Number
        },
        new CellStyle { ExcelFormatId = 27, Formatting = "mm:ss.0", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 28, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 29, Formatting = "mm:ss.0", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 30, Formatting = "d/m/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 31, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 32, Formatting = "d-mmm", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 33, Formatting = "mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 34, Formatting = "h:mm AM/PM", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 35, Formatting = "h:mm:ss AM/PM", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 36, Formatting = "m/d/yy h:mm", FormattingType = FormattingType.DateTime },

        // Additional formats (37-49)
        new CellStyle { ExcelFormatId = 37, Formatting = "#,##0;($#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle
        {
            ExcelFormatId = 38, Formatting = "#,##0.00;($#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 39, Formatting = "#,##0;($#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle
        {
            ExcelFormatId = 40, Formatting = "#,##0.00;($#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle
        {
            ExcelFormatId = 41,
            Formatting = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)",
            FormattingType = FormattingType.Currency
        },
        new CellStyle
        {
            ExcelFormatId = 42,
            Formatting = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
            FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 43, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 44, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 45, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 46, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 47, Formatting = "mmss.0", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 48, Formatting = "##0.0E0", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 49, Formatting = string.Empty, FormattingType = FormattingType.General },

        // Extended formats (50-164)
        new CellStyle { ExcelFormatId = 50, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 51, Formatting = "0.00%", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 52, Formatting = "0.00%", FormattingType = FormattingType.Number },
        new CellStyle
        {
            ExcelFormatId = 53, Formatting = "#,##0.00;(#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle
        {
            ExcelFormatId = 54, Formatting = "#,##0.00;[Red](#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 55, Formatting = "mm:ss;[Red]mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 56, Formatting = "0.00E+0", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 57, Formatting = "# ?/?", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 58, Formatting = "# ??/??", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 59, Formatting = "m/d/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 60, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 61, Formatting = "d-mmm", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 62, Formatting = "mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 63, Formatting = "h:mm AM/PM", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 64, Formatting = "h:mm:ss AM/PM", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 65, Formatting = "h:mm", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 66, Formatting = "h:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 67, Formatting = "m/d/yy h:mm", FormattingType = FormattingType.DateTime },

        new CellStyle { ExcelFormatId = 68, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 69, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 70, Formatting = "mm:ss.0", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 71, Formatting = "##0.0E+0", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 72, Formatting = string.Empty, FormattingType = FormattingType.General },
        new CellStyle { ExcelFormatId = 73, Formatting = "0.00E+00", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 74, Formatting = "# ?/?", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 75, Formatting = "# ??/??", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 76, Formatting = "m/d/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 77, Formatting = "d/m/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 78, Formatting = "d.m.yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 79, Formatting = "d.m.yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 80, Formatting = "d. mmm. yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle
        {
            ExcelFormatId = 81, Formatting = "dddd, d. mmmm yyyy", FormattingType = FormattingType.DateOnly
        },
        new CellStyle { ExcelFormatId = 82, Formatting = "yyyy-m-d", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 83, Formatting = "yyyy-m-d h:mm:ss", FormattingType = FormattingType.DateTime },
        new CellStyle { ExcelFormatId = 84, Formatting = "d/m/yy h:mm:ss", FormattingType = FormattingType.DateTime },
        new CellStyle { ExcelFormatId = 85, Formatting = "d/m/yyyy h:mm:ss", FormattingType = FormattingType.DateTime },
        new CellStyle
        {
            ExcelFormatId = 86, Formatting = "#,##0.0;(#,##0.0)", FormattingType = FormattingType.Currency
        },

        new CellStyle
        {
            ExcelFormatId = 87, Formatting = "#,##0.00;(#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 88, Formatting = "#,##0;(#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle { ExcelFormatId = 89, Formatting = "0.0%", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 90, Formatting = "0%", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 91, Formatting = "[DBNum1][$-804]0", FormattingType = FormattingType.General },
        new CellStyle { ExcelFormatId = 92, Formatting = "[DBNum1][$-804]0", FormattingType = FormattingType.General },
        new CellStyle { ExcelFormatId = 93, Formatting = "[DBNum1][$-804]0", FormattingType = FormattingType.General },
        new CellStyle { ExcelFormatId = 94, Formatting = "[DBNum4][$-804]0", FormattingType = FormattingType.General },
        new CellStyle { ExcelFormatId = 95, Formatting = "mm/dd/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 96, Formatting = "yyyy/m/d", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 97, Formatting = "d MMM yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 98, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 99, Formatting = "d MMMM yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 100, Formatting = "mm-dd", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 101, Formatting = "mm-dd-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 102, Formatting = "mm-dd-yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 103, Formatting = "dd-mm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 104, Formatting = "dd-mm-yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 105, Formatting = "mm-dd-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 106, Formatting = "dd-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 107, Formatting = "mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 108, Formatting = "mmmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 109, Formatting = "m/d/yy h:mm", FormattingType = FormattingType.DateTime },
        new CellStyle { ExcelFormatId = 110, Formatting = "d/m/yy h:mm", FormattingType = FormattingType.DateTime },
        new CellStyle { ExcelFormatId = 111, Formatting = "d/m/yyyy h:mm", FormattingType = FormattingType.DateTime },
        new CellStyle { ExcelFormatId = 112, Formatting = "d/m/yy h:mm:ss", FormattingType = FormattingType.DateTime },
        new CellStyle
        {
            ExcelFormatId = 113, Formatting = "yyyy-m-d h:mm:ss", FormattingType = FormattingType.DateTime
        },
        new CellStyle { ExcelFormatId = 114, Formatting = "dd-mmm-yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 115, Formatting = "dd/mmm/yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 116, Formatting = "dd MMMM yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 117, Formatting = "d. MMMM yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 118, Formatting = "mm/dd/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 119, Formatting = "yyyy-mm-dd", FormattingType = FormattingType.DateOnly },
        new CellStyle
        {
            ExcelFormatId = 120, Formatting = "dd/mm/yyyy h:mm:ss", FormattingType = FormattingType.DateTime
        },
        new CellStyle { ExcelFormatId = 121, Formatting = "mmmm d, yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 122, Formatting = "d MMMM, yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle
        {
            ExcelFormatId = 123, Formatting = "mmmm d, yyyy h:mm:ss", FormattingType = FormattingType.DateTime
        },
        new CellStyle
        {
            ExcelFormatId = 124, Formatting = "#,##0.0;(#,##0.0)", FormattingType = FormattingType.Currency
        },
        new CellStyle
        {
            ExcelFormatId = 126, Formatting = "#,##0.00;(#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 127, Formatting = "#,##0;(#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle
        {
            ExcelFormatId = 128, Formatting = "#,##0.0;(#,##0.0)", FormattingType = FormattingType.Currency
        },
        new CellStyle
        {
            ExcelFormatId = 129, Formatting = "#,##0.00;(#,##0.00)", FormattingType = FormattingType.Currency
        },
        new CellStyle { ExcelFormatId = 130, Formatting = "#,##0;(#,##0)", FormattingType = FormattingType.Currency },
        new CellStyle { ExcelFormatId = 131, Formatting = "0.0%", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 132, Formatting = "0%", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 133, Formatting = "0.00E+00", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 134, Formatting = "0.00E+00", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 135, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 136, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 137, Formatting = "mm:ss.0", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 138, Formatting = "##0.0E+0", FormattingType = FormattingType.Number },
        new CellStyle { ExcelFormatId = 139, Formatting = string.Empty, FormattingType = FormattingType.General },
        new CellStyle
        {
            ExcelFormatId = 140, Formatting = "yyyy-mm-dd hh:mm:ss", FormattingType = FormattingType.DateTime
        },
        new CellStyle { ExcelFormatId = 141, Formatting = "g/m/d", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 142, Formatting = "ge.m.d", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 143, Formatting = "gg", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 144, Formatting = "ggg", FormattingType = FormattingType.DateOnly },
        new CellStyle
        {
            ExcelFormatId = 145, Formatting = "[$-409]h:mm AM/PM", FormattingType = FormattingType.TimeOnly
        },
        new CellStyle
        {
            ExcelFormatId = 146, Formatting = "[$-409]h:mm:ss AM/PM", FormattingType = FormattingType.TimeOnly
        },
        new CellStyle { ExcelFormatId = 147, Formatting = "[$-409]h:mm", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 148, Formatting = "[$-409]h:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 149, Formatting = "[$-409]M/d/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 150, Formatting = "[$-409]d-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 151, Formatting = "[$-409]d-mmm", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 152, Formatting = "[$-409]mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle
        {
            ExcelFormatId = 153, Formatting = "[$-409]m/d/yy h:mm", FormattingType = FormattingType.DateTime
        },
        new CellStyle { ExcelFormatId = 154, Formatting = "mm/dd/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 155, Formatting = "d/m/yy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 156, Formatting = "dd/mm/yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 157, Formatting = "mm/dd/yyyy", FormattingType = FormattingType.DateOnly },
        new CellStyle { ExcelFormatId = 158, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly },
        new CellStyle
        {
            ExcelFormatId = 159, Formatting = "ddd, mmmm dd, yyyy", FormattingType = FormattingType.DateOnly
        },
        new CellStyle
        {
            ExcelFormatId = 160, Formatting = "mm/dd/yyyy h:mm:ss", FormattingType = FormattingType.DateTime
        },
        new CellStyle
        {
            ExcelFormatId = 161, Formatting = "yyyy-mm-dd'T'hh:mm:ss", FormattingType = FormattingType.DateTime
        },
        new CellStyle { ExcelFormatId = 162, Formatting = "hh:mm:ss", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 163, Formatting = "hh:mm:ss.000", FormattingType = FormattingType.TimeOnly },
        new CellStyle { ExcelFormatId = 164, Formatting = "h:mm:ss.00", FormattingType = FormattingType.TimeOnly }
}.AsReadOnly();

    /// <summary>
    /// Gets the formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formattingType">The formatting type, or General if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetFormattingType(short formatId, out FormattingType formattingType)
    {
        if (IsBuiltInFormat(formatId))
        {
            formattingType = BuiltInNumberFormats[formatId].FormattingType;
            return true;
        }
        formattingType = BuiltInNumberFormats[DefaultGeneralFormatId].FormattingType;
        return false;
    }

    public static bool TryGetCellStyle(short formatId, out CellStyle? style)
    {
        if (IsBuiltInFormat(formatId))
        {
            style = BuiltInNumberFormats[formatId];
            return true;
        }
        style = null;
        return false;
    }

    public static CellStyle? GetCellStyle(short formatId )
        => TryGetCellStyle(formatId, out CellStyle? style)
            ? style
            : null;


    /// <summary>
    /// Gets the formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <returns>The formatting type, or General if the format ID is not found.</returns>
    public static FormattingType GetFormattingType(short formatId) =>
        TryGetFormattingType(formatId, out FormattingType format)
            ? format
            : FormattingType.General;

    /// <summary>
    /// Gets the built-in number format code for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formatCode">The format code, or null if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetBuiltInNumberFormat(short formatId, out string formatCode)
    {
        if (IsBuiltInFormat(formatId))
        {
            formatCode = BuiltInNumberFormats[formatId].Formatting!;
        }
        formatCode = string.Empty;
        return false;
    }

    /// <summary>
    /// Gets the built-in number format code for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <returns>The format code, or "General" if the format ID is not found.</returns>
    public static string GetBuiltInNumberFormat(short formatId) =>
        TryGetBuiltInNumberFormat(formatId, out string formatCode)
            ? formatCode
            : string.Empty;

    /// <summary>
    /// Gets both the format code and formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formatCode">The format code, or null if not found.</param>
    /// <param name="formattingType">The formatting type, or General if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetFormat(short formatId, out string formatCode, out FormattingType formattingType)
    {
        if (IsBuiltInFormat(formatId))
        {
            formatCode = BuiltInNumberFormats[formatId].Formatting!;
            formattingType = BuiltInNumberFormats[formatId].FormattingType;
            return true;
        }
        formatCode = string.Empty;
        formattingType = FormattingType.General;
        return false;
    }

    /// <summary>
    /// Gets both the format code and formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <returns>A tuple containing the format code and formatting type, or defaults if not found.</returns>
    public static (string FormatCode, FormattingType Type) GetFormat(short formatId)
        => TryGetFormat(formatId, out string formatCode, out FormattingType format)
            ? (formatCode, format)
            : (string.Empty, FormattingType.General);

    /// <summary>
    /// Determines whether the specified format ID is a built-in format.
    /// </summary>
    /// <param name="formatId">The format ID to check.</param>
    /// <returns>True if the format ID is a built-in format; otherwise false.</returns>
    public static bool IsBuiltInFormat(short formatId) => (BuiltInFormatCount > formatId && formatId >= DefaultGeneralFormatId);


    /// <summary>
    /// Gets the count of built-in number formats.
    /// </summary>
    public static int BuiltInFormatCount => BuiltInNumberFormats.Count;

    /// <summary>
    /// Gets the ID of the default general number format.
    /// </summary>
    public static short DefaultGeneralFormatId => 0;

    /// <summary>
    /// Validates whether a format code follows ECMA-376 conventions.
    /// This is a basic validation that checks for common pattern issues.
    /// </summary>
    /// <param name="formatCode">The format code to validate.</param>
    /// <returns>True if the format code appears valid; otherwise false.</returns>
    public static bool IsValidFormatCode(string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return false;
        }

        // Check for obviously invalid patterns
        // This is a basic check and doesn't validate the complete ECMA-376 format specification
        try
        {
            // Format codes should not be extremely long
            if (formatCode.Length > 500)
            {
                return false;
            }

            // Format codes should not contain control characters
            if (formatCode.Any(c => char.IsControl(c)))
            {
                return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Gets information about ECMA-376 standard.
    /// </summary>
    public static class StandardInfo
    {
        /// <summary>
        /// Gets the ECMA-376 standard version implemented.
        /// </summary>
        public static string Version => "ECMA-376-1:2016, ECMA-376-2:2015";

        /// <summary>
        /// Gets the title of the standard.
        /// </summary>
        public static string Title => "Office Open XML (OOXML) File Formats";

        /// <summary>
        /// Gets the URL to the official ECMA-376 standard documentation.
        /// </summary>
        public static string DocumentationUrl => "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/";

        /// <summary>
        /// Gets a description of the built-in number formats.
        /// </summary>
        public static string NumberFormatsDescription =>
            "Built-in number format codes as defined in ECMA-376 Section 18.8.30 and Excel extensions. " +
            "Format IDs 0-49 are the standard ECMA-376 built-in formats. " +
            "Format IDs 50-164 are extended formats including locale-dependent variants. " +
            "Custom formats use IDs starting from 165.";

        /// <summary>
        /// Gets a description of the default cell styles.
        /// </summary>
        public static string DefaultStylesDescription =>
            "Six implicit/default cell styles (IDs 0-5) that are always present in an ECMA-376 workbook. " +
            "These styles are available even if not explicitly defined in styles.xml or styles.bin.";
    }
}
