using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
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
}
