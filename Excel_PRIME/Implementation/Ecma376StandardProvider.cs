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
    private static readonly Dictionary<short, CellStyle> s_BuiltInNumberFormats = new()
    {
        // https://www.google.com/search?client=firefox-b-d&q=ECMA-376+standard+definitions+for+cell+styling+and+formatting&udm=50&fbs=ADc_l-aN0CWEZBOHjofHoaMMDiKpaEWjvZ2Py1XXV8d8KvlI3p-ML-906rRL_m6h4jR-tdCH-vUIlZq9RzugLEcfjf51b4dfDKizXS4hTwRCZW2TydVcnv1RUVx0SX0axPgL6aA1y5lH4oIQTHc9n3as9K40uq1ucVlSq7hphXixGrVbAHaxl4xbaQRNq-TBoJwkyHSzWgD1m8zRB8KZ0lvZ8gcgw8mFAQ&aep=10&ntc=1&mstk=AUtExfBrrkIu44uA6v1hpUOfIwA1FLcvSwyztg956PFLmym9H3HLadWU4G_XIO0rT-u62dR5h7RXY_GkFuv7v4I4OZL5QNkTRjQIXv8xu4exjOSxpJaFSHsEdcY6V6KqLol4CdAUjOhjs-LNQpumrnSedMNhM1mj9sXpJacczzMVh3ZxNTbk__HZhQVW1jVUwXXSfgFt5TUtMyLOsgEf33rkyP7eUtEyi5g40Yfoo1aEvTyKiVIF3s4mBIm3sCoYqbg64SPXSkl352KtPBXhamjqL4TrYg7-q21oAqMtRvEDpf_3SSIdiZmhrftahdzX-GspJSSxZsuBwT-J4g&aioh=3&csuir=1&mtid=MTQZasfZPInBhbIPiu6M0Q4
        // Six implicit/default cell styles (IDs 0-5) that are always present in an ECMA-376 workbook.
        // These styles are available even if not explicitly defined in styles.xml or styles.bin.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/300280fd-e4fe-4675-a924-4d383af48d3b
        // (Note: IDs 5–8, 23–36, and 41–44 exist in legacy formats or regional variants but are skipped or treated as currency variants in modern standard definitions)
        // The override formats fall into the following ranges:
        // 5 to 8
        // 23 to 26
        // 41 to 44
        // 63 to 66
        // 164 to 392

        // General formats
        { 0, // Default fallback for all cell text/numbers
            new CellStyle { ExcelFormatId = 0, Formatting = "G", FormattingType = FormattingType.General }},

        // Number formats
        { 1, // Integer
            new CellStyle { ExcelFormatId = 1, Formatting = "0", FormattingType = FormattingType.Number } },
        { 2, // Two decimal places
            new CellStyle { ExcelFormatId = 2, Formatting = "0.00", FormattingType = FormattingType.Number } },
        { 3, // Thousands separator
            new CellStyle { ExcelFormatId = 3, Formatting = "#,##0", FormattingType = FormattingType.Number } },
        { 4, // Thousands separator with two decimals
            new CellStyle { ExcelFormatId = 4, Formatting = "#,##0.00", FormattingType = FormattingType.Number } },

        // Currency formats
        { 5, // Currency (No cents, negative in parentheses)
            new CellStyle { ExcelFormatId = 5, Formatting = "$#,##0;($#,##0)", FormattingType = FormattingType.Currency } },
        { 6, // Currency with two decimals
            new CellStyle { ExcelFormatId = 6, Formatting = "$#,##0.00;($#,##0.00)", FormattingType = FormattingType.Currency } },
        { 7, // Currency without decimals (negative in parentheses)
            new CellStyle { ExcelFormatId = 7, Formatting = "$#,##0;($#,##0)", FormattingType = FormattingType.Currency } },
        { 8, // Currency with two decimals (negative in parentheses)
            new CellStyle { ExcelFormatId = 8, Formatting = "$#,##0.00;($#,##0.00)", FormattingType = FormattingType.Currency } },

        // Percentage notation
        { 9, // Percentage integer
            new CellStyle { ExcelFormatId = 9, Formatting = "0%", FormattingType = FormattingType.Percent } },
        { 10, // Percentage with two decimals
            new CellStyle { ExcelFormatId = 10, Formatting = "0.00%", FormattingType = FormattingType.Percent } },

        // Scientific notation
        { 11, // Scientific notation 
            new CellStyle { ExcelFormatId = 11, Formatting = "0.00E+00", FormattingType = FormattingType.Scientific } },

        // Fraction
        { 12, // Single-digit fractions
            new CellStyle { ExcelFormatId = 12, Formatting = "# ?/?", FormattingType = FormattingType.Fraction } },
        { 13, // Two-digit fractions
            new CellStyle { ExcelFormatId = 13, Formatting = "# ??/??", FormattingType = FormattingType.Fraction } },

        // Time/Date formats
        { 14, // Date and time
            new CellStyle { ExcelFormatId = 14, Formatting = "m/d/yyyyy", FormattingType = FormattingType.DateOnly } },
        { 15, // Date only
            new CellStyle { ExcelFormatId = 15, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly } },
        { 16, // Date only
            new CellStyle { ExcelFormatId = 16, Formatting = "d-mmm", FormattingType = FormattingType.DateOnly } },
        { 17, // Date only
            new CellStyle { ExcelFormatId = 17, Formatting = "mmm-yy", FormattingType = FormattingType.DateOnly } },
        { 18, // 12-hour clock
            new CellStyle { ExcelFormatId = 18, Formatting = "h:mm AM/PM", FormattingType = FormattingType.TimeOnly } },
        { 19, // 12-hour clock with seconds
            new CellStyle { ExcelFormatId = 19, Formatting = "h:mm:ss AM/PM", FormattingType = FormattingType.TimeOnly } },
        { 20, // 24-hour clock
            new CellStyle { ExcelFormatId = 20, Formatting = "h:mm", FormattingType = FormattingType.TimeOnly } },
        { 21, // 24-hour clock with seconds
            new CellStyle { ExcelFormatId = 21, Formatting = "h:mm:ss", FormattingType = FormattingType.TimeOnly } },
        { 22, // Date and time
            new CellStyle { ExcelFormatId = 22, Formatting = "m/d/yy h:mm", FormattingType = FormattingType.DateTime } },

        // Additional currency formats (23-36 locale-dependent) or "reserved internal!"
        { 23, // Currency
            new CellStyle { ExcelFormatId = 23, Formatting = "#,##0;(#,##0)", FormattingType = FormattingType.Number } },
        { 24, // Currency with two decimals
            new CellStyle { ExcelFormatId = 24, Formatting = "#,##0.00;(#,##0.00)", FormattingType = FormattingType.Number } },
        { 25, // Currency in red
            new CellStyle { ExcelFormatId = 25, Formatting = "#,##0;[Red](#,##0)", FormattingType = FormattingType.Number } },
        { 26, // Currency with two decimals in red
            new CellStyle { ExcelFormatId = 26, Formatting = "#,##0.00;[Red](#,##0.00)", FormattingType = FormattingType.Number } },
        { 27, // Time only
            new CellStyle { ExcelFormatId = 27, Formatting = "mm:ss.0", FormattingType = FormattingType.TimeOnly } },
        { 28, // Time only
            new CellStyle { ExcelFormatId = 28, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly } },
        { 29, // Time only
            new CellStyle { ExcelFormatId = 29, Formatting = "mm:ss.0", FormattingType = FormattingType.TimeOnly } },
        { 30, // Date only
            new CellStyle { ExcelFormatId = 30, Formatting = "d/m/yy", FormattingType = FormattingType.DateOnly } },
        { 31, // Date only
            new CellStyle { ExcelFormatId = 31, Formatting = "d-mmm-yy", FormattingType = FormattingType.DateOnly } },
        { 32, // Date only
            new CellStyle { ExcelFormatId = 32, Formatting = "d-mmm", FormattingType = FormattingType.DateOnly } },
        { 33, // Date only
            new CellStyle { ExcelFormatId = 33, Formatting = "mmm-yy", FormattingType = FormattingType.DateOnly } },
        { 34, // Time only
            new CellStyle { ExcelFormatId = 34, Formatting = "h:mm AM/PM", FormattingType = FormattingType.TimeOnly } },
        { 35, // Time only
            new CellStyle { ExcelFormatId = 35, Formatting = "h:mm:ss AM/PM", FormattingType = FormattingType.TimeOnly } },
        { 36, // Date and time
            new CellStyle { ExcelFormatId = 36, Formatting = "m/d/yy h:mm", FormattingType = FormattingType.DateTime } },

        // Additional formats (37-49)
        { 37, // Positive / Negative accounting layout
            new CellStyle { ExcelFormatId = 37, Formatting = "#,##0;(#,##0)", FormattingType = FormattingType.Currency } },
        { 38, // Negative values highlighted in red
            new CellStyle { ExcelFormatId = 38, Formatting = "#,##0_);[Red](#,##0)", FormattingType = FormattingType.Currency } },
        { 39, // Accounting with decimals
            new CellStyle { ExcelFormatId = 39, Formatting = "#,##0.00_);(#,##0.00)", FormattingType = FormattingType.Currency } },
        { 40, // Decimals + negative red text
            new CellStyle { ExcelFormatId = 40, Formatting = "#,##0.00_);[Red](#,##0.00)", FormattingType = FormattingType.Currency } },
        { 41, // Standard currency symbol alignment block
            new CellStyle { ExcelFormatId = 41, Formatting = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)", FormattingType = FormattingType.Currency } },
        { 42, // Left-aligned variable accounting whitespace
            new CellStyle { ExcelFormatId = 42, Formatting = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", FormattingType = FormattingType.Currency } },
        { 43, // Currency layout tracking precise cents spacing
            new CellStyle { ExcelFormatId = 43, Formatting = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)", FormattingType = FormattingType.Currency } },
        { 44, // Clean decimal variable aligned currency format
            new CellStyle { ExcelFormatId = 44, Formatting = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", FormattingType = FormattingType.Currency } },

        { 45, // Minutes and seconds
            new CellStyle { ExcelFormatId = 45, Formatting = "mm:ss", FormattingType = FormattingType.TimeOnly } },
        { 46, // Elapsed time (hours can exceed 24)
            new CellStyle { ExcelFormatId = 46, Formatting = "[h]:mm:ss", FormattingType = FormattingType.TimeOnly } },
        { 47, // Split-second duration
            new CellStyle { ExcelFormatId = 47, Formatting = "mmss.0", FormattingType = FormattingType.TimeOnly } },

        { 48, // Alternate scientific notation
            new CellStyle { ExcelFormatId = 48, Formatting = "##0.0E0", FormattingType = FormattingType.Scientific } },
        { 49, // Forces values to render purely as text
            new CellStyle { ExcelFormatId = 49, Formatting = string.Empty, FormattingType = FormattingType.General } },

        // Extended formats (50-164) - These vary by region and cannot be pre-defined!!
    };

    /// <summary>
    /// Gets the formatting type for the specified format ID.
    /// </summary>
    /// <param name="formatId">The format ID to look up.</param>
    /// <param name="formattingType">The formatting type, or General if not found.</param>
    /// <returns>True if the format ID is found; otherwise false.</returns>
    public static bool TryGetFormattingType(short formatId, out FormattingType formattingType)
    {
        if (s_BuiltInNumberFormats.TryGetValue(formatId, out CellStyle? style))
        {
            formattingType = style.FormattingType;
            return true;
        }
        formattingType = FormattingType.General;
        return false;
    }

    public static bool TryGetCellStyle(short formatId, out CellStyle? style)
    {
        if (s_BuiltInNumberFormats.TryGetValue(formatId, out CellStyle? cellStyle))
        {
            style = cellStyle;
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
        if (s_BuiltInNumberFormats.TryGetValue(formatId, out CellStyle? cellStyle))
        {
            formatCode = cellStyle.Formatting!;
            return true;
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
        if (s_BuiltInNumberFormats.TryGetValue(formatId, out CellStyle? cellStyle))
        {
            formatCode = cellStyle.Formatting!;
            formattingType = cellStyle.FormattingType;
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
    public static bool IsBuiltInFormat(short formatId) => s_BuiltInNumberFormats.ContainsKey(formatId);


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

    public static Dictionary<short, CellStyle> GetDefaultStyles()
        => new Dictionary<short, CellStyle>(s_BuiltInNumberFormats);
}
