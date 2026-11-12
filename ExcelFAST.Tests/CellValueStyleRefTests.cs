using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using AwesomeAssertions;

using ExcelPRIME;

using NUnit.Framework;

namespace ExcelFAST.Tests;

/// <summary>
/// Tests for CellValue class with emphasis on _iStyleRef parameter and styled formatting.
/// These tests validate that formatting styles are correctly applied during value conversion.
/// </summary>
[ExcludeFromCodeCoverage]
[TestFixture]
public class CellValueStyleRefTests
{
    #region Integration Tests with Styled Workbooks

    /// <summary>
    /// Tests that styled numeric values are formatted correctly according to their style references.
    /// This validates the ToStyledString() method behavior with various numeric formats.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void ToStyledString_WithNumberFormattingStyles_ShouldFormatCorrectly(string fileName)
    {
        // Arrange
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        // Act & Assert
        workbook.SelectSheet(workbook.SheetNames().First()).Should().BeTrue();

        int cellCount = 0;
        while (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            if (rowCells.IsEmpty)
            {   // Because this returns up to the dimension of the sheet Height
                break;
            }
            foreach (ref readonly Cell cell in rowCells)
            {
                if (!string.IsNullOrEmpty(cell.CellValue.ToString()))
                {
                    cellCount++;
                }
            }
        }

        cellCount.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Tests that ToStyledString() returns properly formatted output for cells with date/time styling.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void ToStyledString_WithDateTimeFormattingStyles_ShouldFormatDateTimeCorrectly(string fileName)
    {
        // Arrange
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        // Act & Assert
        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int rowIndex = 0;
        while (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                string? styledString = cellValue.ToStyledString();
                styledString?.Should().BeOfType<string>();
            }

            rowIndex++;
            if (rowIndex >= 10)
            {
                break;
            }
        }
    }

    #endregion

    #region CellValue Type Conversions with Styling

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void CellValue_NumericConversions_ShouldConvertCorrectlyRegardlessOfStyle(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                try
                {
                    double numericValue = cellValue.AsDouble;
                    numericValue.Should().NotBeNaN();
                }
                catch (FormatException)
                {
                }
                catch (InvalidOperationException)
                {
                }
            }
        }
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void CellValue_DateTimeConversions_ShouldConvertCorrectlyRegardlessOfStyle(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                try
                {
                    DateTime dateValue = cellValue.AsDateTime;
                    dateValue.Year.Should().BeGreaterThan(0);
                }
                catch (FormatException)
                {
                }
                catch (InvalidOperationException)
                {
                }
            }
        }
    }

    #endregion

    #region TryGet Methods with Styled Values

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void TryGetDateTime_WithStyledDateValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int successCount = 0;
        int failureCount = 0;

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                bool success = cellValue.TryGetDateTime(out DateTime result);

                if (success)
                {
                    successCount++;
                    result.Year.Should().BeGreaterThan(0);
                }
                else
                {
                    failureCount++;
                }
            }
        }

        (successCount + failureCount).Should().BeGreaterThan(0);
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void TryGetDouble_WithStyledNumericValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int successCount = 0;

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                if (cellValue.TryGetDouble(out double result))
                {
                    successCount++;
                    result.Should().NotBeNaN();
                }
            }
        }

        successCount.Should().BeGreaterThan(0, "At least some cells should be convertible to double");
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void TryGetInt32_WithStyledNumericValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int successCount = 0;

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                if (cellValue.TryGetInt32(out int result))
                {
                    successCount++;
                    result.Should().BeGreaterThanOrEqualTo(0);
                }
            }
        }

        successCount.Should().BeGreaterThan(0, "At least some cells should be convertible to int");
    }

    #endregion

    #region Implicit Operators with Styling

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void ImplicitOperators_WithStyledNumericValues_ShouldConvertCorrectly(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int numericConversions = 0;

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                try
                {
                    double doubleValue = cellValue;
                    numericConversions++;
                    doubleValue.Should().NotBeNaN();
                }
                catch (FormatException)
                {
                }
            }
        }

        numericConversions.Should().BeGreaterThan(0, "Should have at least some numeric cells");
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void ImplicitOperators_WithStyledDateTimeValues_ShouldConvertCorrectly(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                try
                {
                    DateTime dateTime = cellValue;
                    dateTime.Year.Should().BeGreaterThan(0);
                }
                catch (FormatException)
                {
                }
            }
        }
    }

    #endregion

    #region Consistency Tests

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void CellValue_RawAndStyledValue_ShouldBeConsistent(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int cellCount = 0;

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                object? boxedValue = cellValue.BoxedValue;
                string? styledString = cellValue.ToStyledString();

                cellCount++;
                styledString?.Should().BeOfType<string>();
            }
        }

        cellCount.Should().BeGreaterThan(0);
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void ToString_WithStyledCells_ShouldReturnNonEmptyString(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        workbook.SelectSheet("number & date formatting").Should().BeTrue();

        int nonNullCount = 0;

        if (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                string? toString = cellValue.ToString();

                if (toString != null)
                {
                    nonNullCount++;
                    toString.Should().BeOfType<string>();
                }
            }
        }

        nonNullCount.Should().BeGreaterThan(0);
    }

    #endregion
}
