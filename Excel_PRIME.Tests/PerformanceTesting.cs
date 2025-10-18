using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using ExcelPRIME.Bench;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
internal class PerformanceTesting
{
    [Test]
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx")]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx")]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx")]
    [TestCase("Data/100mb.xlsx")]
    [Explicit("Lot of data being thrown about !")]
    public async Task A010_AccessEveryCellExcel_Prime(string fileName)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        await aecB.AccessEveryCellExcel_Prime();
    }

    [Test]
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx")]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx")]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx")]
    [TestCase("Data/100mb.xlsx")]
    [Explicit("Lot of data being thrown about !")]
    public void A020_AccessEveryCellXlsxHelper(string fileName)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        aecB.AccessEveryCellXlsxHelper();
    }

    [Test]
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx")]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx")]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx")]
    [TestCase("Data/100mb.xlsx")]
    [Explicit("Lot of data being thrown about !")]
    public void A030_AccessEveryCellSylvan(string fileName)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        aecB.AccessEveryCellSylvan();
    }

    [Test]
    [TestCase("Data/sampledocs-50mb-xlsx-file.xlsx")]
    [TestCase("Data/Blank Data 1 Million Rows.xlsx")]
    [TestCase("Data/sampledocs-50mb-xlsx-file-sst.xlsx")]
    [TestCase("Data/100mb.xlsx")]
    [Explicit("Lot of data being thrown about !")]
    public void A040_AccessEveryCellFastExcel(string fileName)
    {
        AccessEveryCellBenchmarks aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        aecB.AccessEveryCellFastExcel();
    }
}
