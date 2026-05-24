using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;


namespace ExcelPRIME.Tests;


[ExcludeFromCodeCoverage]
[NonParallelizable]
[TestFixture]
internal class BugTesting
{
    [Test]
    public async Task Bug_001_SharedStrings()
    {
        const string fileName = "Data/100mb.xlsx";
        int cells = 0;
        using IExcel_PRIMEAsync workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row == null)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                if (row is INullRow)
                {
                    continue;
                }

                IReadOnlyList<ICell?>? rowCells = row.GetAllCells();
                row.Dispose();
                if ( rowCells == null)
                {
                    continue;
                }
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    cells++;
                    if (cells >= 216017)
                    {
                        // When requesting "922473" it should not be returning a blank string!
                        //cell.RawValue.ToString().Should().NotBeNullOrWhiteSpace();
                    }
                }
            }
        }

    }

    [Test]
    public void Bug_020_AsDecimal_FromDecimal()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.Create(val, -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

    [Test]
    public void Bug_020_AsDecimal_FromDouble()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.Create((decimal)(double)val, -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

    [Test]
    public void Bug_020_AsDecimal_FromString()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.Create("41273.28", -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

    [Test]
    public void Bug_020_AsDecimal_FromSpan()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.TryParseOrder("41273.28".AsSpan(), -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

    [Test]
    public async Task Bug_022_EndElement()
    {
        const string fileName = "Data/MissingCells.xlsx";
        int cells = 0;
        using IExcel_PRIMEAsync workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName, new Options{ CellConversionType = CellConversion.ExcelCellType }).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = workbook.GetSheet(sheetName);
            foreach (IRow? row in worksheet!.GetRowData(2))
            {
                IReadOnlyList<ICell?>? rowCells = row.GetAllCells();
                row.Dispose();
                if (rowCells == null)
                {
                    continue;
                }
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    cells++;
                }
            }
        }
        cells.Should().Be(10);
    }

    [Test]
    public async Task Bug_022_EndElement_Async()
    {
        const string fileName = "Data/MissingCells.xlsx";
        int cells = 0;
        using IExcel_PRIMEAsync workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName, new Options { CellConversionType = CellConversion.ExcelCellType }).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            await foreach (IRowAsync? row in worksheet!.GetRowDataAsync(2).ConfigureAwait(false))
            {
                IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                row.Dispose();
                if (rowCells == null)
                {
                    continue;
                }
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    cells++;
                }
            }
        }
        cells.Should().Be(10);

    }

    public static Options[] Option =
    [
        new Options(),
        new Options { ReturnDBNull = true },
        new Options { CellConversionType = CellConversion.ExcelCellType},
        new Options { CellConversionType = CellConversion.ExcelCellType, ReturnDBNull = true }
         // V5 - new Options { CellConversionType = CellConversion.ExcelCellStyle},
         // V5 - new Options { CellConversionType = CellConversion.ExcelCellStyle, ReturnDBNull = true },
         // V5 - new Options { CellConversionType = CellConversion.ForceStyles }
    ];

    [Test]
    [TestCaseSource(nameof(Option))]
    [Explicit]
    public async Task Bug_019_DateInObj(Options options)
    {
        const string fileName = "Data/65K_Records_Data.xlsx";
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName, options).ConfigureAwait(false);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("500000 Sales Records").ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(1, RowCellGet.None).FirstAsync();
        ICell cell = await row.GetCellAsync(6).ConfigureAwait(false)!;
        cell.CellValue?.BoxedValue.Should().BeOfType<DateTime>();//.And.Be(new DateTime(2012, 8, 11)); // Date DD/MM/YYYY
    }

    [CancelAfter(1000)]
    [Test]
    public async Task Bug_027_SkipLines_XLSX(CancellationToken ct)
    {
        const string fileName = "Data/SkipLines.xlsx";
        using IExcel_PRIMEAsync workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName, ct:ct).ConfigureAwait(true);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("SkipLines", ct:ct).ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(3, RowCellGet.None, ct:ct).FirstAsync(ct);
        ICell cell = await row.GetCellAsync(6, ct:ct).ConfigureAwait(false)!;
        cell.CellValue?.AsInt32.Should().Be(1);
    }
    [CancelAfter(1000)]
    [Test]
    public async Task Bug_027_SkipLines_XLSB(CancellationToken ct)
    {
        const string fileName = "Data/SkipLines.xlsb";
        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
        await workbook.OpenAsync(fileName, ct: ct).ConfigureAwait(true);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("SkipLines", ct: ct).ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(3, RowCellGet.None, ct: ct).FirstAsync(ct);
        ICell cell = await row.GetCellAsync(6, ct: ct).ConfigureAwait(false)!;
        cell.CellValue?.AsInt32.Should().Be(1);
    }

    [Test]
    public async Task Bug_028_MultiOpen()
    {
        const string fileName = "Data/100mb.xlsx";
        using IExcel_PRIMEAsync workbook1 = new Excel_PRIME();
        await workbook1.OpenAsync(fileName).ConfigureAwait(true);
        using IExcel_PRIMEAsync workbook2 = new Excel_PRIME();
        await workbook2.OpenAsync(fileName).ConfigureAwait(true);
    }
}
