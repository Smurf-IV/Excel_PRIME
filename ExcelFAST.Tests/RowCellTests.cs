using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using AwesomeAssertions;

using ExcelPRIME;

using NUnit.Framework;


namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
public class RowCellTests
{
    [Test]
    [TestCase("Data/special-char.xlsx")]
    [TestCase("Data/SameKey.xlsx")]
    public void A010_ReadCells(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);
        foreach (string sheetName in workbook.SheetNames())
        {
            workbook.SelectSheet(sheetName);
            while (workbook.GetNextRowData())
            {
                ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
                foreach (ref readonly Cell cell in rowCells)
                {
                    if (cell.CellValue.IsUnknown)
                    {
                        continue;
                    }

                    Console.WriteLine(cell.CellValue.ToString());
                }
            }
        }
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void A020_StyleAndFormattedFile(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

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
        workbook.SelectSheet(workbook.SheetNames().First());
        workbook.Name.Should().Be("text styling");
        int r = 0;
        while (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            if (rowCells.IsEmpty)
            {
                r++;
                continue;
            }

            int c = 0;
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
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

        double?[][] workSheet2Content =
        [
            [123, 2022, 12],
            [123.749273492379, double.NaN/*"Mar – 2022"*/, 12.79879],
            [123.749273492379, 44621, 1232.1],
            [12313.123123123, double.NaN/*"18 mar 22"*/, 123],
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
        workbook.SelectSheet("number & date formatting");
        workbook.Name.Should().Be("number & date formatting");
        r = 0;
        while (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            if (rowCells.IsEmpty)
            {
                r++;
                continue;
            }
            int c = 0;
            foreach (ref readonly Cell cell in rowCells)
            {
                if (c > 0 && cell.CellValue.IsUnknown)
                {
                    break;
                }

                double? expected = workSheet2Content[r][c];
                try
                {
                    if (expected == null)
                    {
                        cell.CellValue.IsUnknown.Should().BeTrue();
                    }
                    else
                    {
                        cell.CellValue.AsDouble.Should().Be(expected);
                    }
                }
                catch
                {
                    if (expected == null || !double.IsNaN(expected.Value))
                    {
                        throw;
                    }
                }

                c++;
            }

            c.Should().Be(workSheet2Content[r].Length);
            r++;
        }

        r.Should().Be(workSheet2Content.Length);
    }

    [Test]
    [TestCase("Data/styledworkbook.xlsx")]
    public void A022_StyleAndFormattedFile_DBNull(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName, new Options { CellConversionType = CellConversion.ExcelCellType, ReturnDBNull = true });

        object[][] workSheet1Content =
        [
            ["a1", "multiline line1\nMultiline line2\nMultiline line 3 multi word", "c1", "d1", "e1"],
            ["bold", "italic", "bold italic", "bold italic underline", DBNull.Value],
            ["bg color1", "bg color and font color", "font color", "text size changed", DBNull.Value],
            ["font changed", "Font + size changed", "<", "&", "'"],
            ["“", "<html>", "<script></script>", "<?xml ?> ", DBNull.Value],
            ["multi format", "\"", " text  ", " t", "t "],
            ["करो हाथों को ऊपर कस आवी गयो", "કેમ છો ", DBNull.Value, DBNull.Value, DBNull.Value]
        ];
        workbook.SheetNames().Should().NotBeEmpty();
        workbook.SelectSheet(workbook.SheetNames().First());
        workbook.Name.Should().Be("text styling");
        int r = 0;
        while (workbook.GetNextRowData())
        {
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            if (rowCells.IsEmpty)
            {
                r++;
                continue;
            }
            int c = 0;
            foreach (ref readonly Cell cell in rowCells)
            {
                if (cell.CellValue.IsUnknown)
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

    [Test]
    public void A030_SkippedRows()
    {
        const string fileName = "Data/solver.xlsx";
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);

        object?[][] workSheet2Content =
        [
            ["Cycle Trader"],
            [],
            [/*null, null, */"Bicycles", "Mopeds", "Child Seats"],
            [/*null, */"Unit Profit", "100", "300", "50"],
            [/*null, null, null, null, null, null, */"Resources", /*null, */"Resources"],
            [/*null, null, null, null, null, null, */"Used", /*null, */"Available"],
            [/*null,*/ "Capital", "300", "1200", "120", /*null,*/ "93000", "≤", "93000"],
            [/*null,*/ "Storage", "0.5", "1", "0.5", /*null,*/ "101", "≤", "101"],
            [],
            [],
            [/*null, null, */"Bicycles", "Mopeds", "Child Seats", /*null, null, null, */"Total Profit"],
            [/*null, */"Order Size", "94", "54", "0", /*null, null, null, */"25600"]
        ];

        workbook.SelectSheet("Solution");
        int r = 0;
        while (workbook.GetNextRowData())
        {
            int c = 0;
            ReadOnlySpan<Cell> rowCells = workbook.GetAllCells();
            if (!rowCells.IsEmpty)
            {
                foreach (ref readonly Cell cell in rowCells)
                {
                    if (cell.CellValue.IsUnknown)
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
                break;
            }
        }
    }

    [Test]
    [TestCase("Data/ValueTest.xlsx")]
    public void A040_ValuesTypesOfCells(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName, new Options { CellConversionType = CellConversion.ForceStyles });
        workbook.SelectSheet("Values");
        workbook.GetNextRowData().Should().BeTrue();

        CellValue cellValue = workbook.GetCellAt(1).Value;
        cellValue.BoxedValue.Should().BeOfType<int>().And.Be(1);
        cellValue.AsInt32.Should().Be(1);

        cellValue = workbook.GetCellAt(2).Value;
        cellValue.BoxedValue.Should().BeOfType<decimal>().And.Be(2.2999999999999998m);
        cellValue.AsDouble.Should().BeApproximately(2.3, 1E-8);

        cellValue = workbook.GetCellAt(3).Value;
        cellValue.BoxedValue.Should().BeOfType<string>().And.Be("abc");
        cellValue.ToString().Should().Be("abc");

        cellValue = workbook.GetCellAt(4).Value;
        cellValue.BoxedValue.Should().BeOfType<bool>().And.Be(true);
        cellValue.AsBoolean.Should().Be(true);

        cellValue = workbook.GetCellAt(5).Value;
        cellValue.BoxedValue.Should().BeOfType<bool>().And.Be(false);
        cellValue.AsBoolean.Should().Be(false);

        cellValue = workbook.GetCellAt(6).Value;
        cellValue.BoxedValue.Should().BeOfType<double>().And.Be(0.01);
        cellValue.AsDouble.Should().Be(0.01);

        cellValue = workbook.GetCellAt(7).Value;
        cellValue.BoxedValue.Should().BeOfType<DateTime>().And.Be(new DateTime(2012, 8, 11));
        cellValue.AsDateTime.Should().Be(new DateTime(2012, 8, 11));

        cellValue = workbook.GetCellAt(8).Value;
        cellValue.BoxedValue.Should().BeOfType<DateTime>().And.Be(new DateTime(2021, 5, 12));
        cellValue.AsDateTime.Should().Be(new DateTime(2021, 5, 12));

        cellValue = workbook.GetCellAt(9).Value;
        cellValue.BoxedValue.Should().BeOfType<DateTime>().And.Be(new DateTime(2011, 5, 23, 19, 12, 30));
        cellValue.AsDateTime.Should().Be(new DateTime(2011, 5, 23, 19, 12, 30));

        cellValue = workbook.GetCellAt(10).Value;
        cellValue.BoxedValue.Should().BeOfType<decimal>().And.Be(2.2999999999999998m);

        cellValue = workbook.GetCellAt(11).Value;
        cellValue.AsDouble.Should().Be(3.3);
    }
}
