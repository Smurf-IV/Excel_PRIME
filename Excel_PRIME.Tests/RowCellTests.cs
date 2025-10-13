using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
public class RowCellTests
{
    [Test]
    [TestCase("Data/special-char.xlsx")]
    [TestCase("Data/SameKey.xlsx")]
    public async Task A010_ReadCells(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(sheetName);
            await foreach (IRow row in worksheet!.GetRowDataAsync())
            {
                await foreach (ICell? cell in row.GetAllCellsAsync())
                {
                    if (cell == null)
                    {   // Because this returns upto the dimension of the sheet width
                        break;
                    }

                    Console.WriteLine(cell.RawValue!.ToString());
                }
            }
        }
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public async Task A020_StyleAndFormattedFile(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);

        string[][] workSheet1Content = new string[][]
    {
        ["a1", "multiline line1\nMultiline line2\nMultiline line 3 multi word", "c1", "d1", "e1"],
        ["bold", "italic", "bold italic", "bold italic underline"],
        ["bg color1", "bg color and font color", "font color", "text size changed"],
        ["font changed", "Font + size changed", "<", "&", "'"],
        ["“", "<html>", "<script></script>", "<?xml ?> "],
        ["multi format" , "\"", " text  ", " t", "t "],
        ["करो हाथों को ऊपर कस आवी गयो", "કેમ છો "]
    };
        workbook.SheetNames().Should().NotBeEmpty();
        using ISheet? worksheet1 = await workbook.GetSheetAsync(workbook.SheetNames().First());
        worksheet1!.Name.Should().Be("text styling");
        int r = 0;
        await foreach (IRow? row in worksheet1.GetRowDataAsync())
        {
            int c = 0;
            await foreach (ICell? cell in row!.GetAllCellsAsync())
            {
                if (cell == null) // Because this returns upto the dimension of the sheet width
                {
                    break;
                }

                cell.RawValue.Should().Be(workSheet1Content[r][c]);
                c++;
            }
            c.Should().Be(workSheet1Content[r].Length);
            r++;
        }
        r.Should().Be(workSheet1Content.Length);

        string?[][] workSheet2Content = new string?[][]
        {
            ["123", "2022", "12"],
            ["123.749273492379", "Mar – 2022", "12.79879"],
            ["123.749273492379", "44621", "1232.1"],
            ["12313.123123123", "18 mar 22", "123"],
            ["13", "200"],
            ["0.00129", "200.90909"],
            ["999.999999", "8980"],
            ["999.999999", "0.508333333333333"],
            [null, "23.3"],
            [null, "1"],
            [null, "2"],
            [null, "2"],
            [null, "-1"],
            [null, "0"],
            [null, "0.5"],
            [null, "0.25"]
        };
        using ISheet? worksheet2 = await workbook.GetSheetAsync("number & date formatting");
        worksheet2!.Name.Should().Be("number & date formatting");
        r = 0;
        await foreach (IRow? row in worksheet2.GetRowDataAsync())
        {
            int c = 0;
            await foreach (ICell? cell in row!.GetAllCellsAsync())
            {
                if (c > 0 && cell == null) // Because this returns upto the dimension of the sheet width
                {
                    break;
                }
                cell?.RawValue.Should().Be(workSheet2Content[r][c]);
                c++;
            }
            c.Should().Be(workSheet2Content[r].Length);
            r++;
        }
        r.Should().Be(workSheet2Content.Length);
    }
}