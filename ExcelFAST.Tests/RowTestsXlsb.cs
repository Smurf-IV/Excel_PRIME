using System;
using System.Diagnostics.CodeAnalysis;
using System.Threading;

using AwesomeAssertions;

using ExcelPRIME;

using NUnit.Framework;

namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class RowTestsXlsb
{
    [Test]
    [TestCase("Data/empty.xlsb")]
    [TestCase("Data/multipleemptysheets.xlsb")]
    public void A010_EmptyXlsb(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsb();
        workbook.Open(fileName);
        workbook.SheetNames().Should().NotBeEmpty();
        foreach (string sheetName in workbook.SheetNames())
        {
            CancellationToken ct = CancellationToken.None;
            workbook.SelectSheet(sheetName, 0, ct);
            while (workbook.GetNextRowData(ct))
            {
                ReadOnlySpan<Cell> cells = workbook.GetAllCells(ct);
            }
        }
    }

    [Test]
    [TestCase("Data/singlesheet.xlsb")]
    [TestCase("Data/styledworkbook.xlsb")]
    [TestCase("Data/verysimple.xlsb")]
    public void A020_CopyBoxedToArray_MatchesCellValues(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsb();
        workbook.Open(fileName);
        foreach (string sheetName in workbook.SheetNames())
        {
            CancellationToken ct = CancellationToken.None;
            workbook.SelectSheet(sheetName, 0, ct);
            while (workbook.GetNextRowData(ct))
            {
                ReadOnlySpan<Cell> cells = workbook.GetAllCells(ct);
                object?[] array = new object?[cells.Length];
                workbook.CopyBoxedToArray(array, ct);

                for (int i = 0; i < cells.Length; i++)
                {
                    array[i].Should().Be(cells[i].CellValue.BoxedValue);
                }

                // Test Span overload with larger buffer
                Span<object?> span = new object?[cells.Length + 5];
                workbook.CopyBoxedToArray(span, ct);
                for (int i = 0; i < cells.Length; i++)
                {
                    span[i].Should().Be(cells[i].CellValue.BoxedValue);
                }
                for (int i = cells.Length; i < span.Length; i++)
                {
                    span[i].Should().BeNull();
                }
            }
        }
    }
}

