using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace Excel_PRIME.Tests;

[ExcludeFromCodeCoverage]
internal class LoadingTests
{
    [Test]
    [TestCase("Data/empty.xlsx")]
    [TestCase("Data/multipleemptysheets.xlsx")]
    public async Task A010_EmptyXlsx(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        workbook.SheetNames().Should().NotBeEmpty();
    }

    [Test]
    [TestCase("Data/invalidfile.xlsx")]
    public async Task A020_NonZipFile(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        Assert.Throws<InvalidDataException>(() => { XlsxReader.OpenWorkbook(path); });
    }

    [Test]
    [TestCase("Data/missingworkbook.xlsx")]
    [TestCase("Data/missingworkbookrelatioship.xlsx")]
    public async Task A030_InvalidXlsx(string fileName)
    {
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName);
        Assert.Throws<XlsxHelperException>(() => { XlsxReader.OpenWorkbook(path); });
    }

    [Test]
    [TestCase("Data/verysimple.xlsx")]
    public async Task A040_DisposeRelasesFile(string fileName)
    {
        using( IExcel_PRIME workbook = new Excel_PRIME())
        {
            await workbook.OpenAsync(fileName);
            //read lock is held
            Assert.Throws<IOException>(() => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
        }
        //no lock, open read/write should work
        using var stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

}
