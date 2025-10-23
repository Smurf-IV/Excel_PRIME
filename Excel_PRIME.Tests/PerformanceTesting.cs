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
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx", 25601276)]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    [TestCase("Data/100mb.xlsx", 8935680)]
    [Explicit("Lot of data being thrown about !")]
    public async Task A010_AccessEveryCellExcel_Prime(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = await aecB.AccessEveryCellExcel_Prime();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx", 25601276)]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    [TestCase("Data/100mb.xlsx", 8935680)]
    [Explicit("Lot of data being thrown about !")]
    public void A020_AccessEveryCellXlsxHelper(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = aecB.AccessEveryCellXlsxHelper();
        cells.Should().BeGreaterThan(expectedCells);
    }

    [Test]
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx", 25601276)]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    [TestCase("Data/100mb.xlsx", 8935680)]
    [Explicit("Lot of data being thrown about !")]
    public async Task A030_AccessEveryCellSylvan(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        int cells = await aecB.AccessEveryCellSylvan();
        cells.Should().BeGreaterThan(expectedCells);
    }

    //[Test]
    //[TestCase("Data/sampledocs-50mb-xlsx-file.xlsx", 7000014)]
    //[TestCase("Data/Blank Data 1 Million Rows.xlsx", 25601276)]
    //[TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx", 7000014)]
    //[TestCase("Data/100mb.xlsx", 8935680)]
    //[Explicit("Lot of data being thrown about !")]
    //public void A040_AccessEveryCellFastExcel(string fileName, int expectedCells)
    //{
    //    AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
    //    int cells = aecB.AccessEveryCellFastExcel();
    //    cells.Should().Be(expectedCells);
    //}
}
