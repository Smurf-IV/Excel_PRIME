using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
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
        cell.CellValue.BoxedValue.Should().BeOfType<DateTime>();//.And.Be(new DateTime(2012, 8, 11)); // Date DD/MM/YYYY
    }

}
