using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using ExcelPRIMEXlsb.Bench;
using ExcelPRIMEXlsb.RangeBench;

using NUnit.Framework;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8629 // Nullable value type may be null.


namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class DefinedRangeTestsXlsb
{
    [Test]
    public async Task A010_ReadNamedRangeAsync()
    {
        const string fileName = "Data/named-range.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        CellValue?[] taxRate = await workbook.GetDefinedRangeAsync("TaxRate").FirstAsync();
        taxRate.First()?.ToString().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");

        // Now do <definedName name="Prices">Sheet1!$A$1:$A$4</definedName>
        CellValue?[][] prices= await workbook.GetDefinedRangeAsync("Prices").ToArrayAsync();
        prices.Should().HaveCount(4);
        prices[0][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(5);
        prices[1][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(4);
        prices[2][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(15);
        prices[3][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(9);
    }

    [Test]
    public void A011_ReadNamedRange()
    {
        const string fileName = "Data/named-range.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        workbook.Open(fileName);
        CellValue?[] taxRate = workbook.GetDefinedRange("TaxRate").First();
        taxRate.First()?.ToString().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");

        // Now do <definedName name="Prices">Sheet1!$A$1:$A$4</definedName>
        CellValue?[][] prices = workbook.GetDefinedRange("Prices").ToArray();
        prices.Should().HaveCount(4);
        prices[0][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(5);
        prices[1][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(4);
        prices[2][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(15);
        prices[3][0]?.BoxedValue.Should().BeOfType<int>().Which.Should().Be(9);
    }

    [Test]
    public async Task A020_DynamicNamedRange()
    {
        const string fileName = "Data/dynamic-named-range.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);

        // Do not fallover with <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
        CellValue?[][] prices = await workbook.GetDefinedRangeAsync("Prices").ToArrayAsync();
        prices.Should().BeNullOrEmpty();
    }

    [Test]
    public async Task A030_LocalSheetID_NamedRange()
    {
        const string fileName = "Data/solver.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName,
            options: new Options { AccessExcelFileInForwardOnlyMode = false }).ConfigureAwait(true);

        // Do not fallover with
        // <definedName name="OrderSize" localSheetId="0">'Try it Yourself'!$C$12:$E$12</definedName>
        // <definedName name="OrderSize">Solution!$C$12:$E$12</definedName>
        CellValue?[] orderSizeS = await workbook.GetDefinedRangeAsync("OrderSize").FirstAsync();
        orderSizeS.Should().HaveCount(3);
        orderSizeS[0]?.BoxedValue.Should().Be(94);
        orderSizeS[1]?.BoxedValue.Should().Be(54);
        orderSizeS[2]?.BoxedValue.Should().Be(0);

        CellValue?[] orderSizeU = await workbook.GetDefinedRangeAsync("OrderSize", "Try it Yourself").FirstAsync();
        orderSizeU.Should().HaveCount(3);
        orderSizeU[0]?.BoxedValue.Should().Be(0);
        orderSizeU[1]?.BoxedValue.Should().Be(0);
        orderSizeU[2]?.BoxedValue.Should().Be(0);

        CellValue?[] orderSizeT = await workbook.GetDefinedRangeAsync("OrderSize (Try it Yourself)").FirstAsync();
        orderSizeT.Should().HaveCount(3);
        orderSizeT[0]?.BoxedValue.Should().Be(0);
        orderSizeT[1]?.BoxedValue.Should().Be(0);
        orderSizeT[2]?.BoxedValue.Should().Be(0);
    }

    [Test]
    [Explicit("Long running tests of external libraries")]
    public void A050_GetRangers100mb()
    {
        const int expected = 1_418_304;
        XlsbRangeBenchmarks aecB = new();
        int cells = aecB.Access100mb(typeof(GRExcelPrimeXlsb));
        cells.Should().Be(expected);
    }

    //[Test]
    //[TestCaseSource(nameof(Rangers))]
    //[Explicit("Other Readers fail!!")]
    //public void A051_GetRangersPivotTable(Type ranger)
    //{
    //    const int expected = 214 * 6 * 2; // Rows * Cols * twice -> Sheet1!$A$1:$F$214
    //    RangeBenchmarks aecB = new();
    //    int cells = aecB.AccessPivotTable(ranger);
    //    cells.Should().Be(expected);
    //}

    [Test]
    public void A052_ExcelPrime_PivotTable()
    {
        const int expected = 214 * 6 * 2; // Rows * Cols * twice -> Sheet1!$A$1:$F$214
        //int cells = aecB.AccessPivotTable(typeof(GetRangeExcelPrime));
        using IGetRangeXlsb getRanger = new GRExcelPrimeXlsb();
        getRanger.LoadFile("Data\\pivot-tables.xlsb");

        //<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">Sheet1!$A$1:$F$214</definedName>
        // In the Xlsb, the `_xlnm.` portion of the name is dropped
        //<definedName name="_FilterDatabase" localSheetId="2" hidden="1">Sheet1!$A$1:$F$214</definedName>
        IEnumerable<IEnumerable<object?>> filterDatabaseSheet = getRanger.GetDefinedRange("_FilterDatabase");
        int cells = filterDatabaseSheet.Sum(row => row.Count());
        IEnumerable<IEnumerable<object?>> filterDatabase = getRanger.GetDefinedRange("_FilterDatabase");
        cells += filterDatabase.Sum(row => row.Count());
        cells.Should().Be(expected);
    }

    [Test]
    public async Task A060_UserRanges()
    {
        const string fileName = "Data/solver.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName,
            options: new Options { AccessExcelFileInForwardOnlyMode = false }).ConfigureAwait(true);

        // Do not fallover with
        Func<Task> sutMethod = async () =>
        {
            await workbook.GetUserRangeAsync("Try it Yourself", "C12:E12").FirstAsync();
        };

        await sutMethod.Should().ThrowAsync<KeyNotFoundException>()
            .WithMessage("* does not exist").ConfigureAwait(false);

        CellValue?[] orderSizeS = await workbook.GetUserRangeAsync("C12:E12", "Solution").FirstAsync();
        orderSizeS.Should().HaveCount(3);
        orderSizeS[0]?.BoxedValue.Should().Be(94);
        orderSizeS[1]?.BoxedValue.Should().Be(54);
        orderSizeS[2]?.BoxedValue.Should().Be(0);

        CellValue?[] orderSizeD = await workbook.GetUserRangeAsync("$C$12:$E$12", "Solution").FirstAsync();
        orderSizeD.Should().HaveCount(3);
        orderSizeD[0]?.BoxedValue.Should().Be(94);
        orderSizeD[1]?.BoxedValue.Should().Be(54);
        orderSizeD[2]?.BoxedValue.Should().Be(0);

        CellValue?[] orderSizeU = await workbook.GetUserRangeAsync("C12", "Solution").FirstAsync();
        orderSizeU.Should().HaveCount(1);
        orderSizeU[0]?.BoxedValue.Should().Be(94);

        CellValue?[] orderSizeC = await workbook.GetUserRangeAsync("$C$12", "Solution").FirstAsync();
        orderSizeC.Should().HaveCount(1);
        orderSizeC[0]?.BoxedValue.Should().Be(94);
    }

}
