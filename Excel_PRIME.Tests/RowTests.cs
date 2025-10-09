using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace Excel_PRIME.Tests;

internal class RowTests
{
    [Test]
    [TestCase("Data/empty.xlsx")]
    [TestCase("Data/multipleemptysheets.xlsx")]
    public async Task A010_EmptyXlsx(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        workbook.SheetNames().Should().NotBeEmpty();
        foreach (var worksheet in workbook.SheetNames())
        {
            using var worksheetReader = worksheet.WorksheetReader;
            foreach (var row in worksheetReader)
            {
            }
        }
    }

}
