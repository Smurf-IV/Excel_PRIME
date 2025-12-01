using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using NUnit.Framework;


namespace ExcelPRIME.Tests;


[ExcludeFromCodeCoverage]
[NonParallelizable]
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

                foreach (ICell? cell in row.GetAllCells())
                {
                    cells++;
                    if (cells >= 216017)
                    {
                        // When requesting "922473" it should not be returning a blank string!
                        //cell.RawValue.ToString().Should().NotBeNullOrWhiteSpace();
                    }
                }
                row.Dispose();
            }
        }

    }
}
