namespace ExcelPRIME.Implementation;

#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
public enum FormattingType
{
    General,
    Number,
    Percent,
    Scientific,
    Fraction,
    Currency,
    DateTime,
    DateOnly,
    TimeOnly
};
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member

/// <summary>
/// Represents the formatting style of a cell in an Excel workbook.
/// </summary>
public class CellStyle
{
    /// <summary>
    /// Gets the number format ID (refers to the numFmt element or built-in format).
    /// </summary>
    public short ExcelFormatId { get; init; }

    /// <summary>
    /// Gets the number format code (e.g., "0.00", "mm/dd/yyyy", "General").
    /// </summary>
    public string? Formatting { get; init; }

    /// <summary>
    /// Gets a value indicating whether the style applies a number format.
    /// </summary>
    public FormattingType FormattingType { get; init; }

    /// <summary>
    /// Gets a string representation of this style.
    /// </summary>
    public override string ToString() =>
        $"Style# Formatting='{Formatting}' (Id={ExcelFormatId})";

    /// <summary>
    /// Gets a value indicating whether FormattingType represents a date or time style.
    /// </summary>
    /// <remarks>
    /// Returns true when FormattingType is one of FormattingType.DateTime, FormattingType.DateOnly,
    /// or FormattingType.TimeOnly.
    /// </remarks>
    public bool IsDateStyle =>
        FormattingType is FormattingType.DateTime or FormattingType.DateOnly or FormattingType.TimeOnly;
}
