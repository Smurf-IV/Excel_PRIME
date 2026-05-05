using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using AwesomeAssertions;

using ExcelPRIME.Bench;

using ExcelPRIMEXlsb.Bench;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class PerformanceTesting
{
    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000012)]
    [TestCase("Blank Data 1 Million Rows.xlsx", 15463982)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000012)]
    [TestCase("100mb.xlsx", 8930256)]
    [Explicit("Lot of data being thrown about !")]
    public void A010_AccessEveryCellExcel_Prime(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName };
        int cells = aecB.Excel_Prime();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsb", 7000012)]
    [TestCase("Blank Data 1 Million Rows.xlsb", 15463982)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsb", 7000012)]
    [TestCase("100mb.xlsb", 8930256)]
    [Explicit("Lot of data being thrown about !")]
    public void A011_AccessEveryCellExcel_Prime_Xlsb(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarksXlsb aecB = new() { FileName = fileName };
        int cells = aecB.Excel_PrimeXlsb();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    [TestCase("Blank Data 1 Million Rows.xlsx", 25601276)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    [TestCase("100mb.xlsx", 8935680)]
    [Explicit("Lot of data being thrown about !")]
    public void A020_AccessEveryCellXlsxHelper(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName };
        int cells = aecB.XlsxHelper();
        cells.Should().BeGreaterThan(expectedCells);
    }

    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    [TestCase("Blank Data 1 Million Rows.xlsx", 25601276)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    [TestCase("100mb.xlsx", 8935680)]
    [Explicit("Lot of data being thrown about !")]
    public async Task A030_AccessEveryCellSylvan(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName };
        int cells = await aecB.SylvanRdr().ConfigureAwait(false);
        cells.Should().BeGreaterThan(expectedCells);
    }

    //[Test]
    //[TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    //[TestCase("Blank Data 1 Million Rows.xlsx", 25601276)]
    //[TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    //[TestCase("100mb.xlsx", 8935680)]
    //[Explicit("Lot of data being thrown about !")]
    //public void A040_AccessEveryCellFastExcel(string fileName, int expectedCells)
    //{
    //    AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
    //    int cells = aecB.FastExcel();
    //    cells.Should().Be(expectedCells);
    //}

    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000012)]
    [TestCase("Blank Data 1 Million Rows.xlsx", 15463982)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000012)]
    [TestCase("100mb.xlsx", 8930256)]
    [Explicit("Lot of data being thrown about !")]
    public async Task A040_ParallelEveryCellAsyncExcel_PrimeTwice(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName };
        int cells = await aecB.PrlAsyncExcel_PrimeTwice().ConfigureAwait(false);
        cells.Should().Be(expectedCells);
    }
}
