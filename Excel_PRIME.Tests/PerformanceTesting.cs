using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using AwesomeAssertions;

using ExcelPRIME.Bench;

using ExcelPRIMEXlsb.Bench;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
[Explicit("Lot of data being thrown about !")]
internal class PerformanceTesting
{
    static readonly object[] TestCases =
        new object[] {
            new object[] {"sampledocs-50mb-xlsx-file", 7000012 },
            new object[] {"Blank Data 1 Million Rows", 15463982 },
            new object[] {"sampledocs-50mb-xlsx-file-sst", 7000012 },
            new object[] {"100mb", 8930256 }
        };

    [Test]
    [TestCaseSource(nameof(TestCases))]
    public void A010_AccessEveryCellExcel_Prime(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName + ".xlsx" };
        int cells = aecB.Excel_Prime();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCaseSource(nameof(TestCases))]
    public void A011_AccessEveryCellExcel_Prime_Xlsb(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarksXlsb aecB = new() { FileName = fileName + ".xlsb" };
        int cells = aecB.Excel_PrimeXlsb();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCaseSource(nameof(TestCases))]
    public void A020_AccessEveryCellXlsxHelper(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName + ".xlsx" };
        int cells = aecB.XlsxHelper();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCaseSource(nameof(TestCases))]
    public async Task A030_AccessEveryCellSylvan(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName + ".xlsx" };
        int cells = await aecB.SylvanRdr().ConfigureAwait(false);
        cells.Should().Be(expectedCells);
    }


    //[Test]
    //[TestCaseSource(nameof(TestCases))]
    //public void A040_AccessEveryCellFastExcel(string fileName, int expectedCells)
    //{
    //    AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName + ".xlsx" };
    //    int cells = aecB.FastExcel();
    //    cells.Should().Be(expectedCells);
    //}

    [Test]
    [TestCaseSource(nameof(TestCases))]
    public async Task A040_ParallelEveryCellAsyncExcel_PrimeTwice(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName + ".xlsx" };
        int cells = await aecB.PrlAsyncExcel_PrimeTwice().ConfigureAwait(false);
        cells.Should().Be(expectedCells);
    }
}
