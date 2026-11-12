using System;
using System.Diagnostics.CodeAnalysis;
using System.Threading;

using AwesomeAssertions;

using ExcelPRIME;
using ExcelPRIME.Bench;

using ExcelPRIMEXlsb.Bench;

using NUnit.Framework;

namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
[Explicit("Large dataset benchmark suite")]
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
    public void A010_AccessEveryCellExcelFastXlsx(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarks aecB = new() { FileName = fileName + ".xlsx" };
        int cells = aecB.ExcelFastXlsx();
        cells.Should().Be(expectedCells);
    }

    [Test]
    [TestCaseSource(nameof(TestCases))]
    public void A011_AccessEveryCellExcelFastXlsb(string fileName, int expectedCells)
    {
        AccessEveryCellBenchmarksXlsb aecB = new() { FileName = fileName + ".xlsb" };
        int cells = aecB.ExcelFastXlsb();
        cells.Should().Be(expectedCells);
    }
}
