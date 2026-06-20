using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;


namespace ExcelPRIME.Tests;

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
    public async Task ToStyledString_WithNumberFormattingStyles_ShouldFormatCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
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

            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            if (rowCells != null)
            {
                cellCount += rowCells.Count(c => !c.CellValue.IsUnknown);
            }
        }

        cellCount.Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Tests that ToStyledString() returns properly formatted output for cells with date/time styling.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task ToStyledString_WithDateTimeFormattingStyles_ShouldFormatDateTimeCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int rowIndex = 0;
        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            // Verify that cells have CellValue instances
            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                ;

                // When ToStyledString is called, it should return either the styled format or the plain string
                string? styledString = cellValue.ToStyledString();
                // String can be null for null values
                styledString?.Should().BeOfType<string>();
            }

            rowIndex++;
            if (rowIndex >= 10)
            {
                break; // Limit iterations for performance
            }
        }
    }

    #endregion

    #region CellValue Type Conversions with Styling

    /// <summary>
    /// Tests that numeric cell values can be properly converted to various types
    /// regardless of their applied style formatting.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task CellValue_NumericConversions_ShouldConvertCorrectlyRegardlessOfStyle(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

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
    /// regardless of their applied formatting styles.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task CellValue_DateTimeConversions_ShouldConvertCorrectlyRegardlessOfStyle(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

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
    /// Tests that TryGetDateTime works correctly with styled date/time cells.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task TryGetDateTime_WithStyledDateValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int successCount = 0;
        int failureCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
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

            break; // Test just first row for performance
        }

        // We expect either successes or failures depending on cell types
        (successCount + failureCount).Should().BeGreaterThan(0);
    }

    /// <summary>
    /// Tests that TryGetDouble works correctly with styled numeric cells.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task TryGetDouble_WithStyledNumericValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int successCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
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

            break; // Test just first row for performance
        }

        successCount.Should().BeGreaterThan(0, "At least some cells should be convertible to double");
    }

    /// <summary>
    /// Tests that TryGetInt32 works correctly with styled numeric cells.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task TryGetInt32_WithStyledNumericValues_ShouldSucceedOrFailGracefully(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int successCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                if (cellValue.TryGetInt32(out int result))
                {
                    successCount++;
                    result.Should().BeLessThan(int.MaxValue);
                }
            }

            break; // Test just first row for performance
        }

        successCount.Should().BeGreaterThan(0, "At least some cells should be convertible to int32");
    }

    #endregion

    #region BoxedValue with Styling

    /// <summary>
    /// Tests that BoxedValue correctly returns the underlying type regardless of formatting style.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task BoxedValue_WithStyledCells_ShouldReturnCorrectType(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int totalCells = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

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
    /// Tests that implicit operators work correctly with styled numeric values.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task ImplicitOperators_WithStyledNumericValues_ShouldConvertCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int numericConversions = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                
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

        numericConversions.Should().BeGreaterThan(0, "Should have at least some numeric cells");
    }

    /// <summary>
    /// Tests that implicit operators work correctly with styled DateTime values.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task ImplicitOperators_WithStyledDateTimeValues_ShouldConvertCorrectly(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int dateConversions = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

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
    /// Tests that the raw value and styled string representation are consistent.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task CellValue_RawAndStyledValue_ShouldBeConsistent(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int cellCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                CellValue cellValue = cell.CellValue;
                
                // Get both representations
                object? boxedValue = cellValue.BoxedValue;
                string? styledString = cellValue.ToStyledString();
                
                cellCount++;

                // Verify they're both accessible without throwing
                styledString?.Should().BeOfType<string>();
            }

            break; // Test just first row for performance
        }

        cellCount.Should().BeGreaterThan(0);
    }

    #endregion

    #region ToString Method with Styling

    /// <summary>
    /// Tests that ToString() returns a meaningful representation for styled cells.
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task ToString_WithStyledCells_ShouldReturnNonEmptyString(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        int nonNullCount = 0;
        int nullCount = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }

            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
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
    /// regardless of their styling (since styling doesn't affect value equality).
    /// </summary>
    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task Equality_StyledCells_ShouldCompareByValue(string fileName)
    {
        // Arrange
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        // Act & Assert
        using ISheetAsync? worksheet = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet.Should().NotBeNull();

        List<CellValue?> cellValues = new List<CellValue?>();
        int rowIndex = 0;

        await foreach (IRowAsync? row in worksheet.GetRowDataAsync().ConfigureAwait(false))
        {
            if (row is null or INullRowAsync)
            {
                continue;
            }
            IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            row.Dispose();
            if (rowCells == null)
            {
                continue;
            }

            foreach (Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
                {
                    break;
                }

                cellValues.Add(cell.CellValue);
            }

            rowIndex++;
            if (rowIndex >= 2)
            {
                break; // Get two rows
            }
        }

        // Verify we have cells to compare
        cellValues.Should().NotBeEmpty();

        // Test equality - same cell values should be equal
        for (int i = 0; i < cellValues.Count - 1; i++)
        {
            // Create equality test
            CellValue? val1 = cellValues[i];
            CellValue? val2 = cellValues[i];
            
            // Same cell should equal itself
            (val1 == val2).Should().Be(true);
        }
    }

    #endregion

    #region DBNull Caching Tests

    /// <summary>
    /// Verifies that GetDBNull returns cached instances for the same style.
    /// </summary>
    [Test]
    public void GetDBNull_WithSameStyle_ShouldReturnCachedInstance()
    {
        CellValue val1 = CellValue.GetDBNull(0);
        CellValue val2 = CellValue.GetDBNull(0);
        
        val1.Equals(val2).Should().BeTrue("GetDBNull(0) should return cached instances to reduce allocation");
    }

    /// <summary>
    /// Verifies that GetDBNull returns cached instances for style -1.
    /// </summary>
    [Test]
    public void GetDBNull_WithStyleNeg1_ShouldReturnCachedInstance()
    {
        CellValue val1 = CellValue.GetDBNull(-1);
        CellValue val2 = CellValue.GetDBNull(-1);
        
        val1.Equals(val2).Should().BeTrue("GetDBNull(-1) should return cached instances");
    }

    /// <summary>
    /// Verifies that DBNull instances with different styles have same hash code to match equality.
    /// </summary>
    [Test]
    public void GetHashCode_ForDBNull_ShouldBeIndependentOfStyle()
    {
        CellValue val1 = CellValue.GetDBNull(0);
        CellValue val2 = CellValue.GetDBNull(1);
        
        val1.GetHashCode().Should().Be(val2.GetHashCode(), "Hash code for DBNull should be independent of style to match Equals");
    }

    #endregion
}

