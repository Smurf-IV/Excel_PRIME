using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;

using AwesomeAssertions;

using NUnit.Framework;


namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class SheetTests
{
    [Test]
    [TestCase("Data/empty.xlsx", 1)]
    [TestCase("Data/multipleemptysheets.xlsx", 3)]
    [TestCase("Data/Hidden.xlsx", 3)]
    public void A010_StepThroughEmptyXlsx(string fileName, int expected)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);
        workbook.SheetNames().Should().HaveCount(expected);
    }

    [Test]
    [TestCase("Data/multisheet1.xlsx", new[] { "one", "two", "three", "b", "a", "a (2)", "a (3)", "b (2)", "one (3)", "two (3)", "three (3)", "b (3)" })]
    [TestCase("Data/singlesheet.xlsx", new[] { "one" })]
    [TestCase("Data/sample_file_bad.xlsx", new[] { "MasterInvoice_Detailed_XLSX" })] // This file contains package relations that use absolute rooted paths instead of relative paths. "/xl/workbook.xml" vs "xl/workbook.xml"
    [TestCase("Data/sample_file_good.xlsx", new[] { "MasterInvoice_Detailed_XLSX" })]
    public void A020_GetsWorkSheets(string fileName, string[] worksheetNames)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);
        workbook.SheetNames().Should().HaveCount(worksheetNames.Length);
        int i = 0;
        foreach (string worksheetName in workbook.SheetNames())
        {
            worksheetName.Should().Be(worksheetNames[i]);
            workbook.SelectSheet(worksheetName).Should().BeTrue(worksheetNames[i]);
            workbook.SheetDimensions.Should().NotBe((-1, -1));

            i++;
        }
    }

    [Test]
    [TestCase("Data/verysimple.xlsx")]
    public void A030_DisposeReleasesFile(string fileName)
    {
        using (IExcelFast workbook = new ExcelFastXlsx())
        {
            workbook.Open(fileName);
            foreach (string worksheetName in workbook.SheetNames())
            {
                workbook.SelectSheet(worksheetName);
                //read lock is held
                Func<FileStream> sutMethod = () => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
                sutMethod.Should().Throw<IOException>();
            }
        }
        //no lock, open read/write should work
        using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

    [Test]
    [TestCase("Data/multisheet1.xlsx")]
    public void A040_ReOpenWorkSheets(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);
        foreach (string worksheetName in workbook.SheetNames())
        {
            workbook.SelectSheet(worksheetName);
        }

        // Now make sure that the sheet source files have not been disposed etc.
        foreach (string worksheetName in workbook.SheetNames())
        {
            workbook.SelectSheet(worksheetName);
        }
    }
}

