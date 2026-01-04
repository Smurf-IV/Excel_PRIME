using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8629 // Nullable value type may be null.

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
public class RowCellTestsXlsb
{
    [Test]
    [TestCase("Data/special-char.xlsb")]
    [TestCase("Data/SameKey.xlsb")]
    public async Task A010_ReadCells(string fileName)
    {
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            await foreach (IRowAsync? row in worksheet!.GetRowDataAsync().ConfigureAwait(false))
            {
                IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    if (cell == null)
                    {
                        // Because this returns upto the dimension of the sheet width
                        break;
                    }

                    Console.WriteLine(cell.CellValue.ToString());
                }
            }
        }
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task A020_StyleAndFormattedFile(string fileName)
    {
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);

        string[][] workSheet1Content =
        [
            ["a1", "multiline line1\nMultiline line2\nMultiline line 3 multi word", "c1", "d1", "e1"],
            ["bold", "italic", "bold italic", "bold italic underline"],
            ["bg color1", "bg color and font color", "font color", "text size changed"],
            ["font changed", "Font + size changed", "<", "&", "'"],
            ["“", "<html>", "<script></script>", "<?xml ?> "],
            ["multi format", "\"", " text  ", " t", "t "],
            ["करो हाथों को ऊपर कस आवी गयो", "કેમ છો "]
        ];
        workbook.SheetNames().Should().NotBeEmpty();
        using ISheetAsync? worksheet1 = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
        worksheet1!.Name.Should().Be("text styling");
        int r = 0;
        await foreach (IRowAsync? row in worksheet1.GetRowDataAsync().ConfigureAwait(false))
        {
            int c = 0;
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            foreach (ICell? cell in rowCells)
            {
                // Because this returns upto the dimension of the sheet width
                if (cell == null) // Because this returns upto the dimension of the sheet width
                {
                    break;
                }

                cell.CellValue.BoxedValue.Should().Be(workSheet1Content[r][c]);
                c++;
            }

            c.Should().Be(workSheet1Content[r].Length);
            r++;
        }

        r.Should().Be(workSheet1Content.Length);

        object?[][] workSheet2Content =
        [
            [123, 2022, 12],
            [123.749273492379, "Mar – 2022", 12.79879],
            [123.749273492379, 44621, 1232.1],
            [12313.123123123, "18 mar 22", 123],
            [13, 200],
            [0.00129, 200.90909],
            [999.999999, 8980],
            [999.999999, 0.508333333333333],
            [null, 23.3],
            [null, 1],
            [null, 2],
            [null, 2],
            [null, -1],
            [null, 0],
            [null, 0.5],
            [null, 0.25]
        ];
        using ISheetAsync? worksheet2 = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
        worksheet2!.Name.Should().Be("number & date formatting");
        r = 0;
        await foreach (IRowAsync? row in worksheet2.GetRowDataAsync().ConfigureAwait(false))
        {
            int c = 0;
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            foreach (ICell? cell in rowCells)
            {
                // Because this returns upto the dimension of the sheet width
                if (c > 0 && cell == null) // Because this returns upto the dimension of the sheet width
                {
                    break;
                }

                cell?.CellValue.BoxedValue.Should().Be(workSheet2Content[r][c]);
                c++;
            }

            c.Should().Be(workSheet2Content[r].Length);
            r++;
        }

        r.Should().Be(workSheet2Content.Length);
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsb")]
    public async Task A021_ParallelStyleAndFormattedFile(string fileName)
    {
        using Excel_PRIMEXlsb workbook1 = new();
        await workbook1.OpenAsync(fileName).ConfigureAwait(false);

        Task.WaitAll(DoSheet1(workbook1), DoSheet2(workbook1));
        return;

        static async Task DoSheet1(IExcel_PRIMEAsync workbook)
        {
            string[][] workSheet1Content =
            [
                ["a1", "multiline line1\nMultiline line2\nMultiline line 3 multi word", "c1", "d1", "e1"],
                ["bold", "italic", "bold italic", "bold italic underline"],
                ["bg color1", "bg color and font color", "font color", "text size changed"],
                ["font changed", "Font + size changed", "<", "&", "'"],
                ["“", "<html>", "<script></script>", "<?xml ?> "],
                ["multi format", "\"", " text  ", " t", "t "],
                ["करो हाथों को ऊपर कस आवी गयो", "કેમ છો "]
            ];
            using ISheetAsync? worksheet1 = await workbook.GetSheetAsync(workbook.SheetNames().First()).ConfigureAwait(false);
            worksheet1!.Name.Should().Be("text styling");
            int r = 0;
            await foreach (IRowAsync? row in worksheet1.GetRowDataAsync().ConfigureAwait(false))
            {
                int c = 0;
                IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    if (cell == null) // Because this returns upto the dimension of the sheet width
                    {
                        break;
                    }

                    cell.CellValue.BoxedValue.Should().Be(workSheet1Content[r][c]);
                    c++;
                }

                c.Should().Be(workSheet1Content[r].Length);
                r++;
            }

            r.Should().Be(workSheet1Content.Length);
        }

        static async Task DoSheet2(IExcel_PRIMEAsync workbook)
        {
            object?[][] workSheet2Content =
            [
                [123, 2022, 12],
                [123.749273492379, "Mar – 2022", 12.79879],
                [123.749273492379, 44621, 1232.1],
                [12313.123123123, "18 mar 22", 123],
                [13, 200], [0.00129, 200.90909],
                [999.999999, 8980],
                [999.999999, 0.508333333333333],
                [null, 23.3],
                [null, 1],
                [null, 2],
                [null, 2],
                [null, -1],
                [null, 0],
                [null, 0.5],
                [null, 0.25]
            ];
            using ISheetAsync? worksheet2 = await workbook.GetSheetAsync("number & date formatting").ConfigureAwait(false);
            worksheet2!.Name.Should().Be("number & date formatting");
            int r = 0;
            await foreach (IRowAsync? row in worksheet2.GetRowDataAsync().ConfigureAwait(false))
            {
                int c = 0;
                IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    if (c > 0 && cell == null) // Because this returns upto the dimension of the sheet width
                    {
                        break;
                    }

                    cell?.CellValue.BoxedValue.Should().Be(workSheet2Content[r][c]);
                    c++;
                }

                c.Should().Be(workSheet2Content[r].Length);
                r++;
            }

            r.Should().Be(workSheet2Content.Length);
        }
    }

    [Test]
    public async Task A030_SkippedRows()
    {
        const string fileName = "Data/solver.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);

        object?[][] workSheet2Content =
        [
            ["Cycle Trader"],
            [],
            [/*null, null, */"Bicycles", "Mopeds", "Child Seats"],
            [/*null, */"Unit Profit", 100, 300, 50],
            [/*null, null, null, null, null, null, */"Resources", /*null, */"Resources"],
            [/*null, null, null, null, null, null, */"Used", /*null, */"Available"],
            [/*null,*/ "Capital", 300, 1200, 120, /*null,*/ 93000, "≤", 93000],
            [/*null,*/ "Storage", 0.5, 1, 0.5, /*null,*/ 101, "≤", 101],
            [],
            [],
            [/*null, null, */"Bicycles", "Mopeds", "Child Seats", /*null, null, null, */"Total Profit"],
            [/*null, */"Order Size", 94, 54, 0, /*null, null, null, */25600]
        ];

        using ISheetAsync? worksheet2 = await workbook.GetSheetAsync("Solution").ConfigureAwait(false);
        int r = 0;
        await foreach (IRowAsync? row in worksheet2.GetRowDataAsync().ConfigureAwait(false))
        {
            int c = 0;
            IReadOnlyList<ICell?>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
            if (rowCells != null)
            {
                foreach (ICell? cell in rowCells)
                {
                    // Because this returns upto the dimension of the sheet width
                    if (cell == null) // Because this returns upto the dimension of the sheet width
                    {
                        continue;
                    }

                    cell.CellValue.BoxedValue.Should().Be(workSheet2Content[r][c]);
                    c++;
                }
            }

            c.Should().Be(workSheet2Content[r].Length);
            r++;
            if (r >= workSheet2Content.Length)
            {
                // This sheet has all sorts of random stuff later on!!
                break;
            }
        }

    }

    [Test]
    [TestCase("Data/ValueTest.xlsb")]
    [Explicit("Types Not implemented yet!")]
    public async Task A040_ValuesTypesOfCells(string fileName)
    {
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("Values").ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(0, RowCellGet.PreGet).FirstAsync();
        ICell? cell = await row.GetCellAsync(1).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<int>().And.Be(1);
        cell = await row.GetCellAsync(2).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<double>().And.Be(2.3);
        cell = await row.GetCellAsync(2).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<string>().And.Be("abc");
        cell = await row.GetCellAsync(3).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<bool>().And.Be(true);
        cell = await row.GetCellAsync(4).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<bool>().And.Be(false);
        cell = await row.GetCellAsync(5).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<double>().And.Be(0.01);//.Within(0.000001); % display
        cell = await row.GetCellAsync(6).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<DateTime>().And.Be(new DateTime(2012, 8, 11)); // Date DD/MM/YYYY
        cell = await row.GetCellAsync(7).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<DateTime>().And.Be(new DateTime(2021, 5, 12));
        cell = await row.GetCellAsync(8).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<DateTime>().And.Be(new DateTime(2011, 5, 23, 19, 12, 30));
        cell = await row.GetCellAsync(9).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<double>().And.Be(2.3);//.Within(0.000001));
        cell = await row.GetCellAsync(10).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<double>().And.Be(3.3);//.Within(0.000001));
        cell = await row.GetCellAsync(11).ConfigureAwait(false);
        cell.CellValue.BoxedValue.Should().BeOfType<string>().And.Be("abcTRUE"); // Number cell type??
    }
}