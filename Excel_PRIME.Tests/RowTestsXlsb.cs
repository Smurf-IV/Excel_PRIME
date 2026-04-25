using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class RowTestsXlsb
{
    [Test]
    [TestCase("Data/empty.xlsb")]
    [TestCase("Data/multipleemptysheets.xlsb")]
    public async Task A010_EmptyXlsb(string fileName)
    {
        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        workbook.SheetNames().Should().NotBeEmpty();
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            await foreach (IRowAsync? row in worksheet!.GetRowDataAsync().ConfigureAwait(false))
            {
                if (row is null or INullRowAsync)
                {
                    continue;
                }
            }
        }
    }

}
