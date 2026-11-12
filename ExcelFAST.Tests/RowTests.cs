using System;
using System.Diagnostics.CodeAnalysis;
using System.Threading;

using AwesomeAssertions;

using ExcelPRIME;

using NUnit.Framework;


namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class RowTests
{
    [Test]
    [TestCase("Data/empty.xlsx")]
    [TestCase("Data/multipleemptysheets.xlsx")]
    public void A010_EmptyXlsx(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
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

}

