using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class SheetTestsXlsb
{
    [Test]
    [TestCase("Data/empty.xlsb", 1)]
    [TestCase("Data/multipleemptysheets.xlsb", 3)]
    [TestCase("Data/Hidden.xlsb", 3)]
    public async Task A010_StepThroughEmptyXlsb(string fileName, int expected)
    {
        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        workbook.SheetNames().Should().HaveCount(expected);
    }

    [Test]
    [TestCase("Data/multisheet1.xlsb", new[] { "one", "two", "three", "b", "a" })]
    [TestCase("Data/singlesheet.xlsb", new[] { "one" })]
    [TestCase("Data/sample_file_bad.xlsb", new[] { "MasterInvoice_Detailed_XLSX" })] // This file contains package relations that use absolute rooted paths instead of relative paths. "/xl/workbook.xml" vs "xl/workbook.xml"
    [TestCase("Data/sample_file_good.xlsb", new[] { "MasterInvoice_Detailed_XLSX" })]
    public async Task A020_GetsWorkSheets(string fileName, string[] worksheetNames)
    {
        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        workbook.SheetNames().Should().HaveCount(worksheetNames.Length);
        int i = 0;
        foreach (string worksheetName in workbook.SheetNames())
        {
            worksheetName.Should().Be(worksheetNames[i]);
            using ISheet? worksheet = await workbook.GetSheetAsync(worksheetName).ConfigureAwait(false);
            worksheet!.Name.Should().Be(worksheetNames[i]);

            i++;
        }
    }

    [Test]
    [TestCase("Data/verysimple.xlsb")]
    public async Task A030_DisposeReleasesFile(string fileName)
    {
        using (IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb())
        {
            await workbook.OpenAsync(fileName).ConfigureAwait(false);
            foreach (string worksheetName in workbook.SheetNames())
            {
                using ISheet? worksheet = await workbook.GetSheetAsync(worksheetName).ConfigureAwait(false);
                //read lock is held
                Func<FileStream> sutMethod = () => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
                sutMethod.Should().Throw<IOException>();
            }
        }
        //no lock, open read/write should work
        using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

    [Test]
    [TestCase("Data/multisheet1.xlsb")]
    public async Task A040_ReOpenWorkSheets(string fileName)
    {
        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        foreach (string worksheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(worksheetName).ConfigureAwait(false);
        }

        // Now make sure that the sheet source files have not been disposed etc.
        foreach (string worksheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(worksheetName).ConfigureAwait(false);
        }
    }
}
