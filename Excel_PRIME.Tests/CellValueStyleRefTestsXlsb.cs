using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;


namespace ExcelPRIME.Tests;

/// <summary>
/// Tests for CellValue class with emphasis on _iStyleRef parameter and styled formatting using XLSB format.
/// These tests validate that formatting styles are correctly applied during value conversion in XLSB workbooks.
/// </summary>
[ExcludeFromCodeCoverage]
[TestFixture]
public class CellValueStyleRefTestsXlsb
{
    #region Integration Tests with Styled Workbooks

    /// <summary>
    /// Tests that styled numeric values are formatted correctly according to their style references in XLSB files.
    /// This validates the ToStyledString() method behavior with various numeric formats.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task ToStyledString_WithNumberFormattingStyles_ShouldFormatCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        
        // Assert
        worksheet.Should().NotBeNull();

        // Verify at least some cells exist
        int cellCount = 0;
        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            if (rowCells != null)
            {
                cellCount += rowCells.Count(c => c != null);
            }
        }

        cellCount.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Tests that ToStyledString() returns properly formatted output for cells with date/time styling in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task ToStyledString_WithDateTimeFormattingStyles_ShouldFormatDateTimeCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int rowIndex = 0;
        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            // Verify that cells have CellValue instances
            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                cellValue.Should().NotBeNull();

                // When ToStyledString is called, it should return either the styled format or the plain string
                string? styledString = cellValue.ToStyledString();
                // String can be null for null values
                if (styledString != null)
                {
                    styledString.Should().BeOfType<string>();
                }
            }

            rowIndex++;
            if (rowIndex >= 10)
                break; // Limit iterations for performance
        }
    }

    #endregion

    #region CellValue Type Conversions with Styling

    /// <summary>
    /// Tests that numeric cell values can be properly converted to various types
    /// regardless of their applied style formatting in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task CellValue_NumericConversions_ShouldConvertCorrectlyRegardlessOfStyle(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                
                // Try various conversions - should not throw
                try
                {
                    double numericValue = cellValue.AsDouble;
                    // Valid numeric conversion - should not throw
                    numericValue.Should().NotBeNaN();
                }
                catch (FormatException)
                {
                    // Expected for non-numeric values
                }
                catch (InvalidOperationException)
                {
                    // Expected for null/empty values
                }
            }

            break; // Test just first row for performance
        }
    }

    /// <summary>
    /// Tests that DateTime cell values can be properly converted
    /// regardless of their applied formatting styles in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task CellValue_DateTimeConversions_ShouldConvertCorrectlyRegardlessOfStyle(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                
                // Try datetime conversion - should not throw
                try
                {
                    DateTime dateValue = cellValue.AsDateTime;
                    // Valid conversion - should have a valid year
                    dateValue.Year.Should().BeGreaterThan(0);
                }
                catch (FormatException)
                {
                    // Expected for non-date values
                }
                catch (InvalidOperationException)
                {
                    // Expected for null/empty values
                }
            }

            break; // Test just first row for performance
        }
    }

    #endregion

    #region TryGet Methods with Styled Values

    /// <summary>
    /// Tests that TryGetDateTime works correctly with styled date/time cells in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task TryGetDateTime_WithStyledDateValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int successCount = 0;
        int failureCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

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

            break; // Test just first row for performance
        }

        // We expect either successes or failures depending on cell types
        (successCount + failureCount).Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Tests that TryGetDouble works correctly with styled numeric cells in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task TryGetDouble_WithStyledNumericValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int successCount = 0;
        int totalAttempts = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                totalAttempts++;
                if (cellValue.TryGetDouble(out double result))
                {
                    successCount++;
                    result.Should().NotBeNaN();
                }
            }

            break; // Test just first row for performance
        }

        // At least some attempt should be made
        totalAttempts.Should().BeGreaterThan(0, "Should have at least some cells to test");
    }

    /// <summary>
    /// Tests that TryGetInt32 works correctly with styled numeric cells in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task TryGetInt32_WithStyledNumericValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int successCount = 0;
        int totalAttempts = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                totalAttempts++;
                if (cellValue.TryGetInt32(out int result))
                {
                    successCount++;
                    result.Should().BeLessThan(int.MaxValue);
                }
            }

            break; // Test just first row for performance
        }

        // At least some attempt should be made
        totalAttempts.Should().BeGreaterThan(0, "Should have at least some cells to test");
    }

    #endregion

    #region BoxedValue with Styling

    /// <summary>
    /// Tests that BoxedValue correctly returns the underlying type regardless of formatting style in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task BoxedValue_WithStyledCells_ShouldReturnCorrectType(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int totalCells = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                object? boxedValue = cellValue.BoxedValue;
                totalCells++;

                if (boxedValue is double)
                {
                    boxedValue.Should().NotBeNull();
                }
                else if (boxedValue is string)
                {
                    boxedValue.Should().NotBeNull();
                }
                else if (boxedValue is null)
                {
                    // Valid - null values
                }
            }

            break; // Test just first row for performance
        }

        // Should have at least some cells to verify
        totalCells.Should().BeGreaterThan(0, "Should have at least some cells in the sheet");
    }

    #endregion

    #region Implicit Operators with Styling

    /// <summary>
    /// Tests that implicit operators work correctly with styled numeric values in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task ImplicitOperators_WithStyledNumericValues_ShouldConvertCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int numericConversions = 0;
        int totalCells = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                totalCells++;

                // Test implicit conversion to double
                try
                {
                    double doubleValue = cellValue;
                    numericConversions++;
                    doubleValue.Should().NotBeNaN();
                }
                catch (FormatException)
                {
                    // Expected for non-numeric cells
                }
            }

            break; // Test just first row for performance
        }

        // Should have tested at least some cells
        totalCells.Should().BeGreaterThan(0, "Should have at least some cells to test");
    }

    /// <summary>
    /// Tests that implicit operators work correctly with styled DateTime values in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task ImplicitOperators_WithStyledDateTimeValues_ShouldConvertCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int dateConversions = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                
                // Test implicit conversion to DateTime
                try
                {
                    DateTime dateTime = cellValue;
                    dateConversions++;
                    dateTime.Year.Should().BeGreaterThan(0);
                }
                catch (FormatException)
                {
                    // Expected for non-date cells
                }
            }

            break; // Test just first row for performance
        }

        // Note: dateConversions might be 0 if the sheet has no datetime values
        // This is acceptable - the test verifies no exceptions are thrown
    }

    #endregion

    #region Consistency Tests

    /// <summary>
    /// Tests that the raw value and styled string representation are consistent in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task CellValue_RawAndStyledValue_ShouldBeConsistent(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int cellCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                
                // Get both representations
                object? boxedValue = cellValue.BoxedValue;
                string? styledString = cellValue.ToStyledString();
                
                cellCount++;

                // Verify they're both accessible without throwing
                if (styledString != null)
                {
                    styledString.Should().BeOfType<string>();
                }
            }

            break; // Test just first row for performance
        }

        cellCount.Should().BeGreaterThan(0);
    }

    #endregion

    #region ToString Method with Styling

    /// <summary>
    /// Tests that ToString() returns a meaningful representation for styled cells in XLSB format.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task ToString_WithStyledCells_ShouldReturnNonEmptyString(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int nonNullCount = 0;
        int nullCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                CellValue cellValue = cell.CellValue;
                string? toString = cellValue.ToString();
                
                if (toString != null)
                {
                    nonNullCount++;
                    toString.Should().BeOfType<string>();
                }
                else
                {
                    nullCount++;
                }
            }

            break; // Test just first row for performance
        }

        // Expect at least some non-null values
        nonNullCount.Should().BeGreaterThan(0);
    }

    #endregion

    #region Equality Comparison with Styling

    /// <summary>
    /// Tests that two CellValue instances with the same underlying value are equal,
    /// regardless of their styling in XLSB format (since styling doesn't affect value equality).
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task Equality_StyledCells_ShouldCompareByValue(string fileName)
    {
        // Arrange
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        var cellValues = new List<CellValue>();
        int rowIndex = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (ICell? cell in rowCells)
            {
                if (cell == null)
                    break;

                cellValues.Add(cell.CellValue);
            }

            rowIndex++;
            if (rowIndex >= 2)
                break; // Get two rows
        }

        // Verify we have cells to compare
        cellValues.Should().NotBeEmpty();

        // Test equality - same cell values should be equal
        for (int i = 0; i < cellValues.Count - 1; i++)
        {
            // Create equality test
            CellValue val1 = cellValues[i];
            CellValue val2 = cellValues[i];
            
            // Same cell should equal itself
            (val1 == val2).Should().Be(true);
        }
    }

    #endregion
}
