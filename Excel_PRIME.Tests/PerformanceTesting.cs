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
    public async Task A010_ReadCells(string fileName)
    {
        var aecB = new AccessEveryCellBenchmarks { FileName = fileName };
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
        var aecB = new AccessEveryCellBenchmarks { FileName = fileName };
        aecB.AccessEveryCellXlsxHelper();
    }
}
