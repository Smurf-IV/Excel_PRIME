using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using ExcelPRIMEXlsb.Bench;
using ExcelPRIMEXlsb.RangeBench;

using NUnit.Framework;


namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
internal class DefinedRangeTestsXlsb
{
    [Test]
    [Explicit]
    public async Task A010_ReadNamedRange()
    {
        const string fileName = "Data/named-range.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        object?[] taxRate = await workbook.GetDefinedRangeAsync("TaxRate").FirstAsync();
        taxRate.FirstOrDefault().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");

        // Now do <definedName name="Prices">Sheet1!$A$1:$A$4</definedName>
        object?[][] prices= await workbook.GetDefinedRangeAsync("Prices").ToArrayAsync();
        prices.Should().HaveCount(4);
        prices[0].Should().BeEquivalentTo(["5"]);
        prices[1].Should().BeEquivalentTo(["4"]);
        prices[2].Should().BeEquivalentTo(["15"]);
        prices[3].Should().BeEquivalentTo(["9"]);
    }

    [Test]
    [Explicit]
    public async Task A020_DynamicNamedRange()
    {
        const string fileName = "Data/dynamic-named-range.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);

        // Do not fallover with <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
        object?[][] prices = await workbook.GetDefinedRangeAsync("Prices").ToArrayAsync();
        prices.Should().BeNullOrEmpty();
    }

    [Test]
    [Explicit]
    public async Task A030_LocalSheetID_NamedRange()
    {
        const string fileName = "Data/solver.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName, 
            options: new Options 
                { 
                    CellConversionType = CellConversion.Number, 
                    AccessExcelFileInForwardOnlyMode = false
                }
            ).ConfigureAwait(true);

        // Do not fallover with
        // <definedName name="OrderSize" localSheetId="0">'Try it Yourself'!$C$12:$E$12</definedName>
        // <definedName name="OrderSize">Solution!$C$12:$E$12</definedName>
        object?[] orderSizeS = await workbook.GetDefinedRangeAsync("OrderSize").FirstAsync();
        orderSizeS.Should().HaveCount(3);
        orderSizeS.Should().HaveElementAt(0, 94);
        orderSizeS.Should().HaveElementAt(1, 54);
        orderSizeS.Should().HaveElementAt(2, 0);

        object?[] orderSizeU = await workbook.GetDefinedRangeAsync("OrderSize", "Try it Yourself").FirstAsync();
        orderSizeU.Should().HaveCount(3);
        orderSizeU.Should().HaveElementAt(0, 0);
        orderSizeU.Should().HaveElementAt(1, 0);
        orderSizeU.Should().HaveElementAt(2, 0);

        object?[] orderSizeT = await workbook.GetDefinedRangeAsync("OrderSize (Try it Yourself)").FirstAsync();
        orderSizeT.Should().HaveCount(3);
        orderSizeT.Should().HaveElementAt(0, 0);
        orderSizeT.Should().HaveElementAt(1, 0);
        orderSizeT.Should().HaveElementAt(2, 0);
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
    [Explicit]
    public void A052_ExcelPrime_PivotTable()
    {
        const int expected = 214 * 6 * 2; // Rows * Cols * twice -> Sheet1!$A$1:$F$214
        //int cells = aecB.AccessPivotTable(typeof(GetRangeExcelPrime));
        using IGetRangeXlsb getRanger = new GRExcelPrimeXlsb();
        getRanger.LoadFile("Data\\pivot-tables.xlsb");

        //<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">Sheet1!$A$1:$F$214</definedName>
        IEnumerable<IEnumerable<object?>> filterDatabaseSheet = getRanger.GetDefinedRange("_xlnm._FilterDatabase", 2);
        int cells = filterDatabaseSheet.Sum(row => row.Count());
        IEnumerable<IEnumerable<object?>> filterDatabase = getRanger.GetDefinedRange("_xlnm._FilterDatabase");
        cells += filterDatabase.Sum(row => row.Count());
        cells.Should().Be(expected);
    }

    [Test]
    [Explicit]
    public async Task A060_UserRanges()
    {
        const string fileName = "Data/solver.xlsb";
        using Excel_PRIMEXlsb workbook = new();
        await workbook.OpenAsync(fileName,
            options: new Options
            {
                CellConversionType = CellConversion.Number,
                AccessExcelFileInForwardOnlyMode = false
            }
        ).ConfigureAwait(true);

        // Do not fallover with
        Func<Task> sutMethod = async () =>
        {
            await workbook.GetUserRangeAsync("Try it Yourself", "C12:E12").FirstAsync();
        };

        await sutMethod.Should().ThrowAsync<KeyNotFoundException>()
            .WithMessage("* does not exist").ConfigureAwait(false);

        object?[] orderSizeS = await workbook.GetUserRangeAsync("C12:E12", "Solution").FirstAsync();
        orderSizeS.Should().HaveCount(3);
        orderSizeS.Should().HaveElementAt(0, 94);
        orderSizeS.Should().HaveElementAt(1, 54);
        orderSizeS.Should().HaveElementAt(2, 0);

        object?[] orderSizeD = await workbook.GetUserRangeAsync("$C$12:$E$12", "Solution").FirstAsync();
        orderSizeD.Should().HaveCount(3);
        orderSizeD.Should().HaveElementAt(0, 94);
        orderSizeD.Should().HaveElementAt(1, 54);
        orderSizeD.Should().HaveElementAt(2, 0);

        object?[] orderSizeU = await workbook.GetUserRangeAsync("C12", "Solution").FirstAsync();
        orderSizeU.Should().HaveCount(1);
        orderSizeU.Should().HaveElementAt(0, 94);

        object?[] orderSizeC = await workbook.GetUserRangeAsync("$C$12", "Solution").FirstAsync();
        orderSizeC.Should().HaveCount(1);
        orderSizeC.Should().HaveElementAt(0, 94);
    }

}
