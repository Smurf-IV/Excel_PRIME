using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading.Tasks;
using System.Xml;

using AwesomeAssertions;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
internal class LoadingTests
{
    [Test]
    [TestCase("test_not_exist.xlsx")]
    public async Task A000_FileNotExist_ThrowsFileNotFoundException(string fileName)
    {
        Func<Task> sutMethod = async () =>
        {
            using IExcel_PRIME workbook = new Excel_PRIME();
            await workbook.OpenAsync(fileName);
        };

        await sutMethod.Should().ThrowAsync<FileNotFoundException>()
            .WithMessage("Could not find file *");
    }

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
        Func<Task> sutMethod = async () =>
        {
            using IExcel_PRIME workbook = new Excel_PRIME();
            await workbook.OpenAsync(fileName);
        };

        await sutMethod.Should().ThrowAsync<InvalidDataException>();
    }

    [Test]
    [TestCase("Data/missingworkbook.xlsx")]
    //[TestCase("Data/missingworkbookrelatioship.xlsx")] This can be loaded by "LibreOffice Calc" ??
    public async Task A030_InvalidXlsx(string fileName)
    {
        Func<Task> sutMethod = async () =>
        {
            using IExcel_PRIME workbook = new Excel_PRIME();
            await workbook.OpenAsync(fileName);
        };
        await sutMethod.Should().ThrowAsync<XmlException> ();
    }

    [Test]
    [TestCase("Data/verysimple.xlsx")]
    public async Task A040_DisposeRelasesFile(string fileName)
    {
        using (IExcel_PRIME workbook = new Excel_PRIME())
        {
            await workbook.OpenAsync(fileName);
            //read lock is held
            Assert.Throws<IOException>(() => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
        }
        //no lock, open read/write should work
        using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

}
