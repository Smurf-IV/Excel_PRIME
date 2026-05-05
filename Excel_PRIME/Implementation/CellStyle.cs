namespace ExcelPRIME.Implementation;

#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
public enum FormattingType
{
    General,
    Number,
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
    /// Gets the index of this style in the styles.xml file.
    /// </summary>
    public int StyleId { get; set; }

    /// <summary>
    /// Gets the number format ID (refers to the numFmt element or built-in format).
    /// </summary>
    public int NumberFormatId { get; set; }

    /// <summary>
    /// Gets the number format code (e.g., "0.00", "mm/dd/yyyy", "General").
    /// </summary>
    public string? NumberFormatCode { get; set; }

    /// <summary>
    /// Gets a value indicating whether the style applies a number format.
    /// </summary>
    public FormattingType FormattingType { get; set; }

    /// <summary>
    /// Gets a string representation of this style.
    /// </summary>
    public override string ToString() =>
        $"Style#{StyleId}: NumberFormat='{NumberFormatCode}' (Id={NumberFormatId})";
}
