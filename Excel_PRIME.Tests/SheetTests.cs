using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace Excel_PRIME.Tests;

internal class SheetTests
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
        }
    }

    [Test]
    [TestCase("Data/multisheet1.xlsx", new[] { "one", "two", "three", "b", "a" })]
    [TestCase("Data/singlesheet.xlsx", new[] { "one" })]
    public async Task A020_GetsWorkSheets(string fileName, string[] worksheetNames)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        workbook.SheetNames().Should().HaveCount(worksheetNames.Length);
        int i = 0;
        foreach (var worksheetName in workbook.SheetNames())
        {
            worksheetName.Should().Be(worksheetNames[i]);
            using var worksheet = await workbook.GetSheetAsync(worksheetName);
            worksheet.Name.Should().Be(worksheetNames[i]);

            i++;
        }
    }

    [Test]
    [TestCase("Data/verysimple.xlsx")]
    public async Task A030_DisposeReleasesFile(string fileName)
    {
        using (IExcel_PRIME workbook = new Excel_PRIME())
        {
            await workbook.OpenAsync(fileName);
            foreach (var worksheetName in workbook.SheetNames())
            {
                using var worksheet = await workbook.GetSheetAsync(worksheetName);
                    //read lock is held
                    Assert.Throws<IOException>(() => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
            }
        }
        //no lock, open read/write should work
        using var stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task A040_StyleAndFormattedFile(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);

        var workSheet1Content = new string[][]
    {
            new string[]{ "a1", "multiline line1\nMultiline line2\nMultiline line 3 multi word", "c1", "d1", "e1" },
            new string[]{ "bold", "italic", "bold italic", "bold italic underline" },
            new string[]{ "bg color1", "bg color and font color", "font color", "text size changed" },
            new string[]{ "font changed", "Font + size changed", "<", "&", "'" },
            new string[]{ "“", "<html>", "<script></script>", "<?xml ?> " },
            new string[]{ "multi format" , "\"", " text  ", " t", "t " },
            new string[]{ "करो हाथों को ऊपर कस आवी गयो", "કેમ છો " }
    };
        Assert.NotEmpty(workbook.Worksheets);
        var worksheet1 = workbook.Worksheets.First();
        Assert.Equal("text styling", worksheet1.Name);
        var r = 0;
        foreach (var row in worksheet1.WorksheetReader)
        {
            int c = 0;
            for (; c < row.Cells.Length; c++)
            {
                Assert.Equal(workSheet1Content[r][c], row.Cells[c].CellValue);
            }
            Assert.Equal(workSheet1Content[r].Length, c);
            r++;
        }
        Assert.Equal(workSheet1Content.Length, r);

        var workSheet2Content = new string[][]
        {
            new string[]{ "123", "2022", "12"  },
            new string[]{ "123.749273492379", "Mar – 2022", "12.79879" },
            new string[]{ "123.749273492379", "44621", "1232.1" },
            new string[]{ "12313.123123123", "18 mar 22", "123" },
            new string[]{ "13", "200" },
            new string[]{ "0.00129", "200.90909" },
            new string[]{ "999.999999", "8980" },
            new string[]{ "999.999999", "0.508333333333333" },
            new string[]{ "23.3" },
            new string[]{ "1" },
            new string[]{ "2" },
            new string[]{ "2" },
            new string[]{ "-1"},
            new string[]{ "0" },
            new string[]{ "0.5" },
            new string[]{ "0.25" }
        };
        var worksheet2 = workbook.Worksheets.ElementAt(1);
        Assert.Equal("number & date formatting", worksheet2.Name);
        r = 0;
        foreach (var row in worksheet2.WorksheetReader)
        {
            int c = 0;
            for (; c < row.Cells.Length; c++)
            {
                Assert.Equal(workSheet2Content[r][c], row.Cells[c].CellValue);
            }
            Assert.Equal(workSheet2Content[r].Length, c);
            r++;
        }
        Assert.Equal(workSheet2Content.Length, r);
    }
}
