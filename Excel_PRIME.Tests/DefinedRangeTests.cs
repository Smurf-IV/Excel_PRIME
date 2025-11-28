using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using ExcelPRIME.RangeBench;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
internal class DefinedRangeTests
{
    [Test]
    public async Task A010_ReadNamedRange()
    {
        const string fileName = "Data/named-range.xlsx";
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        object?[] taxRate = await workbook.GetDefinedRangeAsync("TaxRate").FirstAsync();
        taxRate.FirstOrDefault().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");

        // Now do <definedName name="Prices">Sheet1!$A$1:$A$4</definedName>
        var prices= await workbook.GetDefinedRangeAsync("Prices").ToArrayAsync();
        prices.Should().HaveCount(4);
        prices[0].Should().BeEquivalentTo(["5"]);
        prices[1].Should().BeEquivalentTo(["4"]);
        prices[2].Should().BeEquivalentTo(["15"]);
        prices[3].Should().BeEquivalentTo(["9"]);
    }

    [Test]
    public async Task A020_DynamicNamedRange()
    {
        const string fileName = "Data/dynamic-named-range.xlsx";
        using Excel_PRIME workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);

        // Do not fallover with <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
        var prices = await workbook.GetDefinedRangeAsync("Prices").ToArrayAsync();
    }

    [Test]
    public async Task A030_LocalSheetID_NamedRange()
    {
        const string fileName = "Data/solver.xlsx";
        using Excel_PRIME workbook = new();
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
        var orderSizeS = await workbook.GetDefinedRangeAsync("OrderSize").FirstAsync();
        orderSizeS.Should().HaveCount(3);
        orderSizeS.Should().HaveElementAt(0, 94);
        orderSizeS.Should().HaveElementAt(1, 54);
        orderSizeS.Should().HaveElementAt(2, 0);

        var orderSizeU = await workbook.GetDefinedRangeAsync("OrderSize", "Try it Yourself").FirstAsync();
        orderSizeU.Should().HaveCount(3);
        orderSizeU.Should().HaveElementAt(0, 0);
        orderSizeU.Should().HaveElementAt(1, 0);
        orderSizeU.Should().HaveElementAt(2, 0);

        var orderSizeT = await workbook.GetDefinedRangeAsync("OrderSize (Try it Yourself)").FirstAsync();
        orderSizeT.Should().HaveCount(3);
        orderSizeT.Should().HaveElementAt(0, 0);
        orderSizeT.Should().HaveElementAt(1, 0);
        orderSizeT.Should().HaveElementAt(2, 0);
    }

    public static Type[] Rangers = // for multiple arguments it's an IEnumerable of IGetRange's
    [
        typeof(GetRangeExcelPrime),  //  8.8 
        typeof(GetRangeClosedXML),  // 44.1
        typeof(GetRangeEPPlus),   // 15.2 ->  V8.3 | 14.1 V7.3.2
        typeof(GetRangeFreeSpire),   // 26.1
        typeof(GetRangeAsposeCells)
    ];

    [Test]
    [TestCaseSource(nameof(Rangers))]
    public void A050_GetRangers(Type ranger)
    {
        const int expected = 1_418_304;
        RangeBenchmarks aecB = new();
        int cells = aecB.Access100mb(ranger);
        cells.Should().Be(expected);
    }

}
