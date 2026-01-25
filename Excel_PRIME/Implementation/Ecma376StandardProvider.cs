using System.Collections.Generic;
using System.Linq;

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
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string FormatCode, FormattingType Type)> BuiltInNumberFormats = new Dictionary<int, (string, FormattingType)>
    {
        // General formats
        { 0, ("General", FormattingType.General) },
        
        // Number formats
        { 1, ("0", FormattingType.Number) },
        { 2, ("0.00", FormattingType.Number) },
        { 3, ("#,##0", FormattingType.Number) },
        { 4, ("#,##0.00", FormattingType.Number) },
        
        // Currency formats
        { 5, ("$#,##0;($#,##0)", FormattingType.Currency) },
        { 6, ("$#,##0.00;($#,##0.00)", FormattingType.Currency) },
        { 7, ("$#,##0;($#,##0)", FormattingType.Currency) },
        { 8, ("$#,##0.00;($#,##0.00)", FormattingType.Currency) },
        
        // Scientific notation
        { 9, ("0.00E+00", FormattingType.Number) },
        { 10, ("#,##0.0", FormattingType.Number) },
        
        // Text format
        { 11, ("@", FormattingType.General) },
        
        // Time/Date formats
        { 12, ("mm:ss", FormattingType.TimeOnly) },
        { 13, ("[h]:mm:ss", FormattingType.TimeOnly) },
        { 14, ("mm/dd/yyyy", FormattingType.DateTime) },
        { 15, ("d-mmm-yy", FormattingType.DateOnly) },
        { 16, ("d-mmm", FormattingType.DateOnly) },
        { 17, ("mmm-yy", FormattingType.DateOnly) },
        { 18, ("h:mm AM/PM", FormattingType.TimeOnly) },
        { 19, ("h:mm:ss AM/PM", FormattingType.TimeOnly) },
        { 20, ("h:mm", FormattingType.TimeOnly) },
        { 21, ("h:mm:ss", FormattingType.TimeOnly) },
        { 22, ("m/d/yy h:mm", FormattingType.DateTime) },
        
        // Additional currency formats (23-36 locale-dependent)
        { 23, ("#,##0;(#,##0)", FormattingType.Number) },
        { 24, ("#,##0.00;(#,##0.00)", FormattingType.Number) },
        { 25, ("#,##0;[Red](#,##0)", FormattingType.Number) },
        { 26, ("#,##0.00;[Red](#,##0.00)", FormattingType.Number) },
        { 27, ("mm:ss.0", FormattingType.TimeOnly) },
        { 28, ("[h]:mm:ss", FormattingType.TimeOnly) },
        { 29, ("mm:ss.0", FormattingType.TimeOnly) },
        { 30, ("d/m/yy", FormattingType.DateOnly) },
        { 31, ("d-mmm-yy", FormattingType.DateOnly) },
        { 32, ("d-mmm", FormattingType.DateOnly) },
        { 33, ("mmm-yy", FormattingType.DateOnly) },
        { 34, ("h:mm AM/PM", FormattingType.TimeOnly) },
        { 35, ("h:mm:ss AM/PM", FormattingType.TimeOnly) },
        { 36, ("m/d/yy h:mm", FormattingType.DateTime) },
        
        // Additional formats (37-49)
        { 37, ("#,##0;($#,##0)", FormattingType.Currency) },
        { 38, ("#,##0.00;($#,##0.00)", FormattingType.Currency) },
        { 39, ("#,##0;($#,##0)", FormattingType.Currency) },
        { 40, ("#,##0.00;($#,##0.00)", FormattingType.Currency) },
        { 41, ("_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)", FormattingType.Currency) },
        { 42, ("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", FormattingType.Currency) },
        { 43, ("mm:ss", FormattingType.TimeOnly) },
        { 44, ("[h]:mm:ss", FormattingType.TimeOnly) },
        { 45, ("mm:ss", FormattingType.TimeOnly) },
        { 46, ("[h]:mm:ss", FormattingType.TimeOnly) },
        { 47, ("mmss.0", FormattingType.TimeOnly) },
        { 48, ("##0.0E0", FormattingType.Number) },
        { 49, ("@", FormattingType.General) },
        
        // Extended formats (50-164)
        { 50, ("mm:ss", FormattingType.TimeOnly) },
        { 51, ("0.00%", FormattingType.Number) },
        { 52, ("0.00%", FormattingType.Number) },
        { 53, ("#,##0.00;(#,##0.00)", FormattingType.Currency) },
        { 54, ("#,##0.00;[Red](#,##0.00)", FormattingType.Currency) },
        { 55, ("mm:ss;[Red]mm:ss", FormattingType.TimeOnly) },
        { 56, ("0.00E+0", FormattingType.Number) },
        { 57, ("# ?/?", FormattingType.Number) },
        { 58, ("# ??/??", FormattingType.Number) },
        { 59, ("m/d/yy", FormattingType.DateOnly) },
        { 60, ("d-mmm-yy", FormattingType.DateOnly) },
        { 61, ("d-mmm", FormattingType.DateOnly) },
        { 62, ("mmm-yy", FormattingType.DateOnly) },
        { 63, ("h:mm AM/PM", FormattingType.TimeOnly) },
        { 64, ("h:mm:ss AM/PM", FormattingType.TimeOnly) },
        { 65, ("h:mm", FormattingType.TimeOnly) },
        { 66, ("h:mm:ss", FormattingType.TimeOnly) },
        { 67, ("m/d/yy h:mm", FormattingType.DateTime) },
        { 68, ("mm:ss", FormattingType.TimeOnly) },
        { 69, ("[h]:mm:ss", FormattingType.TimeOnly) },
        { 70, ("mm:ss.0", FormattingType.TimeOnly) },
        { 71, ("##0.0E+0", FormattingType.Number) },
        { 72, ("@", FormattingType.General) },
        { 73, ("0.00E+00", FormattingType.Number) },
        { 74, ("# ?/?", FormattingType.Number) },
        { 75, ("# ??/??", FormattingType.Number) },
        { 76, ("m/d/yy", FormattingType.DateOnly) },
        { 77, ("d/m/yy", FormattingType.DateOnly) },
        { 78, ("d.m.yy", FormattingType.DateOnly) },
        { 79, ("d.m.yyyy", FormattingType.DateOnly) },
        { 80, ("d. mmm. yyyy", FormattingType.DateOnly) },
        { 81, ("dddd, d. mmmm yyyy", FormattingType.DateOnly) },
        { 82, ("yyyy-m-d", FormattingType.DateOnly) },
        { 83, ("yyyy-m-d h:mm:ss", FormattingType.DateTime) },
        { 84, ("d/m/yy h:mm:ss", FormattingType.DateTime) },
        { 85, ("d/m/yyyy h:mm:ss", FormattingType.DateTime) },
        { 86, ("#,##0.0;(#,##0.0)", FormattingType.Currency) },
        { 87, ("#,##0.00;(#,##0.00)", FormattingType.Currency) },
        { 88, ("#,##0;(#,##0)", FormattingType.Currency) },
        { 89, ("0.0%", FormattingType.Number) },
        { 90, ("0%", FormattingType.Number) },
        { 91, ("[DBNum1][$-804]0", FormattingType.General) },
        { 92, ("[DBNum1][$-804]0", FormattingType.General) },
        { 93, ("[DBNum1][$-804]0", FormattingType.General) },
        { 94, ("[DBNum4][$-804]0", FormattingType.General) },
        { 95, ("mm/dd/yy", FormattingType.DateOnly) },
        { 96, ("yyyy/m/d", FormattingType.DateOnly) },
        { 97, ("d MMM yy", FormattingType.DateOnly) },
        { 98, ("d-mmm-yy", FormattingType.DateOnly) },
        { 99, ("d MMMM yy", FormattingType.DateOnly) },
        { 100, ("mm-dd", FormattingType.DateOnly) },
        { 101, ("mm-dd-yy", FormattingType.DateOnly) },
        { 102, ("mm-dd-yyyy", FormattingType.DateOnly) },
        { 103, ("dd-mm-yy", FormattingType.DateOnly) },
        { 104, ("dd-mm-yyyy", FormattingType.DateOnly) },
        { 105, ("mm-dd-yy", FormattingType.DateOnly) },
        { 106, ("dd-mmm-yy", FormattingType.DateOnly) },
        { 107, ("mmm-yy", FormattingType.DateOnly) },
        { 108, ("mmmm-yy", FormattingType.DateOnly) },
        { 109, ("m/d/yy h:mm", FormattingType.DateTime) },
        { 110, ("d/m/yy h:mm", FormattingType.DateTime) },
        { 111, ("d/m/yyyy h:mm", FormattingType.DateTime) },
        { 112, ("d/m/yy h:mm:ss", FormattingType.DateTime) },
        { 113, ("yyyy-m-d h:mm:ss", FormattingType.DateTime) },
        { 114, ("dd-mmm-yyyy", FormattingType.DateOnly) },
        { 115, ("dd/mmm/yyyy", FormattingType.DateOnly) },
        { 116, ("dd MMMM yyyy", FormattingType.DateOnly) },
        { 117, ("d. MMMM yyyy", FormattingType.DateOnly) },
        { 118, ("mm/dd/yy", FormattingType.DateOnly) },
        { 119, ("yyyy-mm-dd", FormattingType.DateOnly) },
        { 120, ("dd/mm/yyyy h:mm:ss", FormattingType.DateTime) },
        { 121, ("mmmm d, yyyy", FormattingType.DateOnly) },
        { 122, ("d MMMM, yyyy", FormattingType.DateOnly) },
        { 123, ("mmmm d, yyyy h:mm:ss", FormattingType.DateTime) },
        { 124, ("mm/dd/yyyy", FormattingType.DateOnly) },
        { 125, ("#,##0.0;(#,##0.0)", FormattingType.Currency) },
        { 126, ("#,##0.00;(#,##0.00)", FormattingType.Currency) },
        { 127, ("#,##0;(#,##0)", FormattingType.Currency) },
        { 128, ("#,##0.0;(#,##0.0)", FormattingType.Currency) },
        { 129, ("#,##0.00;(#,##0.00)", FormattingType.Currency) },
        { 130, ("#,##0;(#,##0)", FormattingType.Currency) },
        { 131, ("0.0%", FormattingType.Number) },
        { 132, ("0%", FormattingType.Number) },
        { 133, ("0.00E+00", FormattingType.Number) },
        { 134, ("0.00E+00", FormattingType.Number) },
        { 135, ("mm:ss", FormattingType.TimeOnly) },
        { 136, ("[h]:mm:ss", FormattingType.TimeOnly) },
        { 137, ("mm:ss.0", FormattingType.TimeOnly) },
        { 138, ("##0.0E+0", FormattingType.Number) },
        { 139, ("@", FormattingType.General) },
        { 140, ("yyyy-mm-dd hh:mm:ss", FormattingType.DateTime) },
        { 141, ("g/m/d", FormattingType.DateOnly) },
        { 142, ("ge.m.d", FormattingType.DateOnly) },
        { 143, ("gg", FormattingType.DateOnly) },
        { 144, ("ggg", FormattingType.DateOnly) },
        { 145, ("[$-409]h:mm AM/PM", FormattingType.TimeOnly) },
        { 146, ("[$-409]h:mm:ss AM/PM", FormattingType.TimeOnly) },
        { 147, ("[$-409]h:mm", FormattingType.TimeOnly) },
        { 148, ("[$-409]h:mm:ss", FormattingType.TimeOnly) },
        { 149, ("[$-409]M/d/yy", FormattingType.DateOnly) },
        { 150, ("[$-409]d-mmm-yy", FormattingType.DateOnly) },
        { 151, ("[$-409]d-mmm", FormattingType.DateOnly) },
        { 152, ("[$-409]mmm-yy", FormattingType.DateOnly) },
        { 153, ("[$-409]m/d/yy h:mm", FormattingType.DateTime) },
        { 154, ("mm/dd/yy", FormattingType.DateOnly) },
        { 155, ("d/m/yy", FormattingType.DateOnly) },
        { 156, ("dd/mm/yyyy", FormattingType.DateOnly) },
        { 157, ("mm/dd/yyyy", FormattingType.DateOnly) },
        { 158, ("d-mmm-yy", FormattingType.DateOnly) },
        { 159, ("ddd, mmmm dd, yyyy", FormattingType.DateOnly) },
        { 160, ("mm/dd/yyyy h:mm:ss", FormattingType.DateTime) },
        { 161, ("yyyy-mm-dd'T'hh:mm:ss", FormattingType.DateTime) },
        { 162, ("hh:mm:ss", FormattingType.TimeOnly) },
        { 163, ("hh:mm:ss.000", FormattingType.TimeOnly) },
        { 164, ("h:mm:ss.00", FormattingType.TimeOnly) },
    }.AsReadOnly();

    /// <summary>
    /// Gets the formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formattingType">The formatting type, or General if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetFormattingType(int formatId, out FormattingType formattingType)
    {
        if (BuiltInNumberFormats.TryGetValue(formatId, out (string _, FormattingType Type) format))
        {
            formattingType = format.Type;
            return true;
        }
        formattingType = FormattingType.General;
        return false;
    }

    /// <summary>
    /// Gets the formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <returns>The formatting type, or General if the format ID is not found.</returns>
    public static FormattingType GetFormattingType(int formatId) =>
        BuiltInNumberFormats.TryGetValue(formatId, out (string FormatCode, FormattingType Type) format)
            ? format.Type
            : FormattingType.General;

    /// <summary>
    /// Gets the built-in number format code for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formatCode">The format code, or null if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetBuiltInNumberFormat(int formatId, out string? formatCode)
    {
        if (BuiltInNumberFormats.TryGetValue(formatId, out (string FormatCode, FormattingType _) format))
        {
            formatCode = format.FormatCode;
            return true;
        }
        formatCode = null;
        return false;
    }

    /// <summary>
    /// Gets the built-in number format code for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <returns>The format code, or "General" if the format ID is not found.</returns>
    public static string GetBuiltInNumberFormat(int formatId) =>
        BuiltInNumberFormats.TryGetValue(formatId, out (string FormatCode, FormattingType Type) format)
            ? format.FormatCode
            : "General";

    /// <summary>
    /// Gets both the format code and formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formatCode">The format code, or null if not found.</param>
    /// <param name="formattingType">The formatting type, or General if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetFormat(int formatId, out string? formatCode, out FormattingType formattingType)
    {
        if (BuiltInNumberFormats.TryGetValue(formatId, out (string FormatCode, FormattingType Type) format))
        {
            formatCode = format.FormatCode;
            formattingType = format.Type;
            return true;
        }
        formatCode = null;
        formattingType = FormattingType.General;
        return false;
    }

    /// <summary>
    /// Gets both the format code and formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <returns>A tuple containing the format code and formatting type, or defaults if not found.</returns>
    public static (string FormatCode, FormattingType Type) GetFormat(int formatId)
    {
        if (BuiltInNumberFormats.TryGetValue(formatId, out (string FormatCode, FormattingType Type) format))
        {
            return format;
        }
        return ("General", FormattingType.General);
    }

    /// <summary>
    /// Default cell styles that are implicitly present in every ECMA-376 workbook.
    /// These styles are always available, even if not explicitly defined in styles.xml or styles.bin.
    /// 
    /// Reference: ECMA-376-1:2016 Section 18.8.10 (Cell Formats - cellXfs)
    /// </summary>
    public static class DefaultStyles
    {
        /// <summary>
        /// Gets all default styles as a dictionary.
        /// </summary>
        /// <returns>A dictionary of default style IDs to CellStyle objects.</returns>
        public static IReadOnlyDictionary<int, CellStyle> GetAll() =>
            new Dictionary<int, CellStyle>
            {
                { 0, CreateNormalStyle() },
                { 1, CreateCommaStyle() },
                { 2, CreateCommaDecimalStyle() },
                { 3, CreateCurrencyStyle() },
                { 4, CreateCurrencyDecimalStyle() },
                { 5, CreatePercentStyle() },
            }.AsReadOnly();

        /// <summary>
        /// Creates Style 0 - Normal/General
        /// The implicit default style applied to all cells unless otherwise specified.
        /// </summary>
        private static CellStyle CreateNormalStyle() => new()
        {
            StyleId = 0,
            NumberFormatId = 0,
            NumberFormatCode = "General",
            FormattingType = FormattingType.General,
        };

        /// <summary>
        /// Creates Style 1 - Comma format
        /// Number with thousands separator (e.g., 1,234)
        /// </summary>
        private static CellStyle CreateCommaStyle() => new()
        {
            StyleId = 1,
            NumberFormatId = 3,
            NumberFormatCode = "#,##0",
            FormattingType = FormattingType.Number,
        };

        /// <summary>
        /// Creates Style 2 - Comma (2 decimal places)
        /// Number with thousands separator and 2 decimal places (e.g., 1,234.56)
        /// </summary>
        private static CellStyle CreateCommaDecimalStyle() => new()
        {
            StyleId = 2,
            NumberFormatId = 4,
            NumberFormatCode = "#,##0.00",
            FormattingType = FormattingType.Number,
        };

        /// <summary>
        /// Creates Style 3 - Currency
        /// Currency format with thousands separator (e.g., $1,234; ($1,234))
        /// </summary>
        private static CellStyle CreateCurrencyStyle() => new()
        {
            StyleId = 3,
            NumberFormatId = 5,
            NumberFormatCode = "$#,##0;($#,##0)",
            FormattingType = FormattingType.Currency,
        };

        /// <summary>
        /// Creates Style 4 - Currency (2 decimal places)
        /// Currency format with thousands separator and 2 decimal places (e.g., $1,234.56; ($1,234.56))
        /// </summary>
        private static CellStyle CreateCurrencyDecimalStyle() => new()
        {
            StyleId = 4,
            NumberFormatId = 6,
            NumberFormatCode = "$#,##0.00;($#,##0.00)",
            FormattingType = FormattingType.Currency,
        };

        /// <summary>
        /// Creates Style 5 - Percent
        /// Percentage format (e.g., 25%)
        /// </summary>
        private static CellStyle CreatePercentStyle() => new()
        {
            StyleId = 5,
            NumberFormatId = 7,
            NumberFormatCode = "0%",
            FormattingType = FormattingType.Number,
        };
    }

    /// <summary>
    /// Determines whether the specified format ID is a built-in format.
    /// </summary>
    /// <param name="formatId">The format ID to check.</param>
    /// <returns>True if the format ID is a built-in format; otherwise false.</returns>
    public static bool IsBuiltInFormat(int formatId) => BuiltInNumberFormats.ContainsKey(formatId);

    /// <summary>
    /// Gets all built-in number format IDs.
    /// </summary>
    /// <returns>A collection of format IDs that are built-in.</returns>
    public static IEnumerable<int> GetBuiltInFormatIds() => BuiltInNumberFormats.Keys;

    /// <summary>
    /// Gets the count of built-in number formats.
    /// </summary>
    public static int BuiltInFormatCount => BuiltInNumberFormats.Count;

    /// <summary>
    /// Gets the default number format code for "General" format.
    /// </summary>
    public static string DefaultGeneralFormat => "General";

    /// <summary>
    /// Gets the ID of the default general number format.
    /// </summary>
    public static int DefaultGeneralFormatId => 0;

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
