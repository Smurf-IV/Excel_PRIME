using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using AwesomeAssertions;

using ExcelFAST.Implementation;

using ExcelPRIME;

using ExcelPRIMEXlsb.Bench;
using ExcelPRIMEXlsb.RangeBench;

using NUnit.Framework;

namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class DefinedRangeTestsXlsb
{
    [Test]
    public void A010_ReadNamedRange()
    {
        const string fileName = "Data/named-range.xlsb";
        using IExcelFast workbook = new ExcelFastXlsb();
        workbook.Open(fileName);
        PooledCellView[] taxRate = workbook.GetDefinedRange("TaxRate").First();
        taxRate.First().ToString().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");

        // Now do <definedName name="Prices">Sheet1!$A$1:$A$4</definedName>
        PooledCellView[][] prices = workbook.GetDefinedRange("Prices").ToArray();
        prices.Should().HaveCount(4);
        prices[0][0].Value.BoxedValue.Should().BeOfType<int>().Which.Should().Be(5);
        prices[1][0].Value.BoxedValue.Should().BeOfType<int>().Which.Should().Be(4);
        prices[2][0].Value.BoxedValue.Should().BeOfType<int>().Which.Should().Be(15);
        prices[3][0].Value.BoxedValue.Should().BeOfType<int>().Which.Should().Be(9);
    }

    [Test]
    public void A020_DynamicNamedRange()
    {
        const string fileName = "Data/dynamic-named-range.xlsb";
        using IExcelFast workbook = new ExcelFastXlsb();
        workbook.Open(fileName);

        // Do not fallover with <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
        PooledCellView[][] prices = workbook.GetDefinedRange("Prices").ToArray();
        prices.Should().BeNullOrEmpty();
    }

    [Test]
    public void A030_LocalSheetID_NamedRange()
    {
        const string fileName = "Data/solver.xlsb";
        using IExcelFast workbook = new ExcelFastXlsb();
        workbook.Open(fileName, options: new Options { AccessExcelFileInForwardOnlyMode = false });

        // Do not fallover with
        // <definedName name="OrderSize" localSheetId="0">'Try it Yourself'!$C$12:$E$12</definedName>
        // <definedName name="OrderSize">Solution!$C$12:$E$12</definedName>
        PooledCellView[] orderSizeS = workbook.GetDefinedRange("OrderSize").First();
        orderSizeS.Should().HaveCount(3);
        orderSizeS[0].Value.BoxedValue.Should().Be(94);
        orderSizeS[1].Value.BoxedValue.Should().Be(54);
        orderSizeS[2].Value.BoxedValue.Should().Be(0);

        PooledCellView[] orderSizeU = workbook.GetDefinedRange("OrderSize", "Try it Yourself").First();
        orderSizeU.Should().HaveCount(3);
        orderSizeU[0].Value.BoxedValue.Should().Be(0);
        orderSizeU[1].Value.BoxedValue.Should().Be(0);
        orderSizeU[2].Value.BoxedValue.Should().Be(0);

        PooledCellView[] orderSizeT = workbook.GetDefinedRange("OrderSize (Try it Yourself)").First();
        orderSizeT.Should().HaveCount(3);
        orderSizeT[0].Value.BoxedValue.Should().Be(0);
        orderSizeT[1].Value.BoxedValue.Should().Be(0);
        orderSizeT[2].Value.BoxedValue.Should().Be(0);
    }

    [Test]
    [Explicit("Long running tests of external libraries")]
    public void A050_GetRangers100mb()
    {
        const int expected = 1_418_304;
        XlsbRangeBenchmarks aecB = new();
        int cells = aecB.Access100mb(typeof(GRExcelFastXlsb));
        cells.Should().Be(expected);
    }

    [Test]
    public void A052_ExcelFastXlsb_PivotTable()
    {
        const int expected = 214 * 6 * 2; // Rows * Cols * twice -> Sheet1!$A$1:$F$214
        using IGetRangeXlsb getRanger = new GRExcelFastXlsb();
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
    public void A060_UserRanges()
    {
        const string fileName = "Data/solver.xlsb";
        using IExcelFast workbook = new ExcelFastXlsb();
        workbook.Open(fileName, options: new Options { AccessExcelFileInForwardOnlyMode = false });

        Action sutMethod = () =>
        {
            workbook.GetUserRange("Try it Yourself", "C12:E12").First();
        };

        sutMethod.Should().Throw<KeyNotFoundException>()
            .WithMessage("* does not exist");

        PooledCellView[] orderSizeS = workbook.GetUserRange("C12:E12", "Solution").First();
        orderSizeS.Should().HaveCount(3);
        orderSizeS[0].Value.BoxedValue.Should().Be(94);
        orderSizeS[1].Value.BoxedValue.Should().Be(54);
        orderSizeS[2].Value.BoxedValue.Should().Be(0);

        PooledCellView[] orderSizeD = workbook.GetUserRange("$C$12:$E$12", "Solution").First();
        orderSizeD.Should().HaveCount(3);
        orderSizeD[0].Value.BoxedValue.Should().Be(94);
        orderSizeD[1].Value.BoxedValue.Should().Be(54);
        orderSizeD[2].Value.BoxedValue.Should().Be(0);

        PooledCellView[] orderSizeU = workbook.GetUserRange("C12", "Solution").First();
        orderSizeU.Should().HaveCount(1);
        orderSizeU[0].Value.BoxedValue.Should().Be(94);

        PooledCellView[] orderSizeC = workbook.GetUserRange("$C$12", "Solution").First();
        orderSizeC.Should().HaveCount(1);
        orderSizeC[0].Value.BoxedValue.Should().Be(94);
    }
}
