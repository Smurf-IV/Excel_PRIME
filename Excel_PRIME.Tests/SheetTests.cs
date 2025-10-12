using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

using AwesomeAssertions;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

internal class SheetTests
{
    [Test]
    [TestCase("Data/empty.xlsx", 1)]
    [TestCase("Data/multipleemptysheets.xlsx", 3)]
    public async Task A010_StepThroughEmptyXlsx(string fileName, int expected)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        workbook.SheetNames().Should().HaveCount(expected);
    }

    [Test]
    [TestCase("Data/multisheet1.xlsx", new[] { "one", "two", "three", "b", "a" })]
    [TestCase("Data/singlesheet.xlsx", new[] { "one" })]
    public async Task A020_GetsWorkSheets(string fileName, string[] worksheetNames)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        workbook.SheetNames().Should().HaveCount(worksheetNames.Length);
        int i = 0;
        foreach (string worksheetName in workbook.SheetNames())
        {
            worksheetName.Should().Be(worksheetNames[i]);
            using ISheet? worksheet = await workbook.GetSheetAsync(worksheetName);
            worksheet!.Name.Should().Be(worksheetNames[i]);

            i++;
        }
    }

    [Test]
    [TestCase("Data/verysimple.xlsx")]
    public async Task A030_DisposeReleasesFile(string fileName)
    {
        using (IExcel_PRIME workbook = new Excel_PRIME())
        {
            await workbook.OpenAsync(fileName);
            foreach (string worksheetName in workbook.SheetNames())
            {
                using ISheet? worksheet = await workbook.GetSheetAsync(worksheetName);
                //read lock is held
                Func<FileStream> sutMethod = () => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
                sutMethod.Should().Throw<IOException>();
            }
        }
        //no lock, open read/write should work
        using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

}
