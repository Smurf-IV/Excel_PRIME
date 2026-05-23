using System;
using System.Globalization;
using NUnit.Framework;
using AwesomeAssertions;
using ExcelPRIME;

namespace Excel_PRIME.Tests;

[TestFixture]
public class CellValueDecimalTests
{
    [Test]
    public void CellValue_CreateFromDecimal_ShouldStoreDecimal()
    {
        // Arrange
        decimal expected = 123.456m;

        // Act
        CellValue cellValue = CellValue.Create(expected, -1);

        // Assert
        cellValue.BoxedValue.Should().Be(expected);
        cellValue.AsDecimal.Should().Be(expected);
        cellValue.AsDouble.Should().Be((double)expected);
        cellValue.ToString().Should().Be(expected.ToString(CultureInfo.InvariantCulture));
    }

    [Test]
    public void CellValue_DecimalPrecision_ShouldBePreserved()
    {
        // Arrange
        // A value that cannot be represented exactly by double
        decimal expected = 0.0000000000000000000000000001m;

        // Act
        CellValue cellValue = CellValue.Create(expected, -1);

        // Assert
        cellValue.AsDecimal.Should().Be(expected);
        // double would likely lose precision or underflow
        cellValue.AsDouble.Should().Be((double)expected);
    }

    [Test]
    public void CellValue_DecimalEquality_ShouldWork()
    {
        // Arrange
        decimal val1 = 123.45m;
        decimal val2 = 123.45m;
        decimal val3 = 456.78m;

        CellValue cell1 = CellValue.Create(val1, -1);
        CellValue cell2 = CellValue.Create(val2, -1);
        CellValue cell3 = CellValue.Create(val3, -1);

        // Assert
        cell1.Equals(cell2).Should().BeTrue();
        cell1.Equals(cell3).Should().BeFalse();
        cell1.GetHashCode().Should().Be(cell2.GetHashCode());
    }

    [Test]
    public void CellValue_AsDecimal_FromNumeric_ShouldWork()
    {
        // Arrange
        decimal val = 123.45m;
        CellValue cellValue = CellValue.Create(val, -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

    [Test]
    public void CellValue_AsDecimal_FromString_ShouldWork()
    {
        // Arrange
        string val = "123.45";
        CellValue cellValue = CellValue.Create(val, -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(123.45m);
    }
    
    [Test]
    public void CellValue_TryFormat_Decimal_ShouldWork()
    {
        // Arrange
        decimal val = 123.456m;
        CellValue cellValue = CellValue.Create(val, -1);
        Span<char> buffer = stackalloc char[32];

        // Act
        bool success = cellValue.TryFormat(buffer, out int charsWritten, default, null);

        // Assert
        success.Should().BeTrue();
        buffer.Slice(0, charsWritten).ToString().Should().Be(val.ToString(CultureInfo.InvariantCulture));
    }
}
