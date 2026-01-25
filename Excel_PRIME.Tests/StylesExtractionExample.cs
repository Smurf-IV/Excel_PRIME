using System;
using System.Collections.Generic;
using System.Threading;

using ExcelPRIME.Implementation;

namespace ExcelPRIME.Examples;

/// <summary>
/// Example demonstrating how to extract and use cell styles from XLSX files.
/// </summary>
public class StylesExtractionExample
{
    /// <summary>
    /// Demonstrates extracting styles from an Excel workbook and using them to format cell values.
    /// </summary>
    public static void ExtractAndUseStyles()
    {
        CancellationToken ct = CancellationToken.None;

        // Open an Excel file
        using (Excel_PRIME excel = new Excel_PRIME())
        {
            excel.Open("Sample.xlsx");

            // Get a sheet
            using (ISheet? sheet = excel.GetSheet("Sheet1", ct: ct))
            {
                if (sheet == null)
                {
                    Console.WriteLine("Sheet not found");
                    return;
                }

                // Iterate through rows
                foreach (IRow? row in sheet.GetRowData(ct: ct))
                {
                    if (row == null)
                    {
                        continue;
                    }

                    IReadOnlyList<ICell?>? cells = row.GetAllCells(ct);
                    if (cells == null)
                    {
                        continue;
                    }

                    foreach (ICell? cell in cells)
                    {
                        if (cell?.CellValue == null)
                        {
                            continue;
                        }

                        // Example 1: Get the basic string representation
                        string basicString = cell.CellValue.ToString() ?? "null";
                        Console.WriteLine($"Cell value: {basicString}");

                        // Example 2: Get the styled string representation using extracted styles
                        // The styles are automatically extracted when the workbook is opened
                        string styledString = cell.CellValue.ToStyledString() ?? "null";
                        Console.WriteLine($"Styled cell value: {styledString}");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Demonstrates directly accessing cell styles from the workbook.
    /// </summary>
    public static void AccessCellStylesDirectly()
    {
        CancellationToken ct = CancellationToken.None;

        using (Excel_PRIME excel = new Excel_PRIME())
        {
            excel.Open("Sample.xlsx");

            // After opening, styles are automatically extracted and available in the instance context
            // You can access them programmatically if needed

            // Example: Check available styles
            using (ISheet? sheet = excel.GetSheet("Sheet1", ct: ct))
            {
                if (sheet == null)
                {
                    return;
                }

                foreach (IRow? row in sheet.GetRowData(ct: ct))
                {
                    if (row == null)
                    {
                        continue;
                    }

                    IReadOnlyList<ICell?>? cells = row.GetAllCells(ct);
                    if (cells == null)
                    {
                        continue;
                    }

                    foreach (ICell? cell in cells)
                    {
                        if (cell?.CellValue == null)
                        {
                            continue;
                        }

                        // Use ToStyledString with available styles
                        // Note: You may need to pass styles through your application's context
                        string value = cell.CellValue.ToString() ?? "";
                        Console.WriteLine($"Value: {value}");
                    }
                }
            }
        }
    }
}

/// <summary>
/// Example of working with number formats from cell styles.
/// </summary>
public class NumberFormatExample
{
    public static void DemonstrateNumberFormats()
    {
        CancellationToken ct = CancellationToken.None;

        using (Excel_PRIME excel = new Excel_PRIME())
        {
            excel.Open("Sample.xlsx");

            using (ISheet? sheet = excel.GetSheet("Sheet1", ct: ct))
            {
                if (sheet == null)
                {
                    return;
                }

                foreach (IRow? row in sheet.GetRowData(ct: ct))
                {
                    if (row == null)
                    {
                        continue;
                    }

                    IReadOnlyList<ICell?>? cells = row.GetAllCells(ct);
                    if (cells == null)
                    {
                        continue;
                    }

                    foreach (ICell? cell in cells)
                    {
                        if (cell?.CellValue == null)
                        {
                            continue;
                        }

                        CellValue cellValue = cell.CellValue;

                        // For numeric values, the style's number format can be applied
                        try
                        {
                            double numValue = cellValue.AsDouble;
                            
                            // Common formats
                            string formatted = numValue switch
                            {
                                _ when numValue >= 1000 => numValue.ToString("N2"), // 1,234.56
                                _ when numValue < 1 && numValue > 0 => numValue.ToString("P"), // 50.00%
                                _ => numValue.ToString("G") // General format
                            };

                            Console.WriteLine($"Number: {numValue}, Formatted: {formatted}");
                        }
                        catch
                        {
                            // Value is not numeric, skip
                        }
                    }
                }
            }
        }
    }
}

/// <summary>
/// Helper methods for accessing CellStyle information.
/// </summary>
public static class CellStyleExtensions
{
    /// <summary>
    /// Checks if a cell value has a numeric format applied.
    /// </summary>
    public static bool HasNumericFormat(this CellStyle style) => !string.IsNullOrEmpty(style?.NumberFormatCode);

    /// <summary>
    /// Checks if a cell value has a date/time format applied.
    /// </summary>
    public static bool HasDateTimeFormat(this CellStyle style)
    {
        if (string.IsNullOrEmpty(style?.NumberFormatCode))
        {
            return false;
        }

        var span = style.NumberFormatCode.AsSpan();
        return span.ContainsAny("dmyhs");
    }

    /// <summary>
    /// Checks if a cell value has a percentage format applied.
    /// </summary>
    public static bool HasPercentageFormat(this CellStyle style) => style?.NumberFormatCode?.Contains('%') == true;

    /// <summary>
    /// Gets the number of decimal places for a numeric format.
    /// </summary>
    public static int GetDecimalPlaces(this CellStyle style)
    {
        if (!string.IsNullOrEmpty(style?.NumberFormatCode))
        {
            return 0;
        }

        int count = 0;
        bool foundDecimal = false;

        foreach (char c in style.NumberFormatCode)
        {
            if (c == '.')
            {
                foundDecimal = true;
            }
            else if (foundDecimal && c == '0')
            {
                count++;
            }
            else if (foundDecimal && c != '0')
            {
                break;
            }
        }

        return count;
    }
}
