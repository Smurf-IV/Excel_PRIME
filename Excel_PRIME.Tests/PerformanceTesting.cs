using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using AwesomeAssertions;

using ExcelPRIME.Bench;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
internal class PerformanceTesting
{
    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000012)]
    [TestCase("Blank Data 1 Million Rows.xlsx", 14314945)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000012)]
    [TestCase("100mb.xlsx", 8930256)]
    [Explicit("Lot of data being thrown about !")]
    public async Task A010_AccessEveryCellExcel_Prime(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = await aecB.AccessEveryCellExcel_Prime().ConfigureAwait(false);
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
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = aecB.AccessEveryCellXlsxHelper();
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
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = await aecB.AccessEveryCellSylvan().ConfigureAwait(false);
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
    //    int cells = aecB.AccessEveryCellFastExcel();
    //    cells.Should().Be(expectedCells);
    //}

    [Test]
    [TestCase("sampledocs-50mb-xlsx-file.xlsx", 7000012)]
    [TestCase("Blank Data 1 Million Rows.xlsx", 14314945)]
    [TestCase("sampledocs-50mb-xlsx-file-sst.xlsx", 7000012)]
    [TestCase("100mb.xlsx", 8930256)]
    [Explicit("Lot of data being thrown about !")]
    public async Task A040_ParallelEveryCellAsyncExcel_PrimeTwice(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = await aecB.ParallelEveryCellAsyncExcel_PrimeTwice().ConfigureAwait(false);
        cells.Should().Be(expectedCells);
    }
}
