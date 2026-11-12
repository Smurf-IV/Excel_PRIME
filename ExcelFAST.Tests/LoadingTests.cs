using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;

using AwesomeAssertions;

using NUnit.Framework;


namespace ExcelFAST.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class LoadingTests
{
    [Test]
    [TestCase("test_not_exist.xlsx")]
    public void A000_FileNotExist_ThrowsFileNotFoundException(string fileName)
    {
        Action sutMethod =  () =>
        {
            using IExcelFast workbook = new ExcelFastXlsx();
            workbook.Open(fileName);
        };

        sutMethod.Should().Throw<FileNotFoundException>()
            .WithMessage("Could not find file *");
    }

    [Test]
    [TestCase("Data/empty.xlsx")]
    [TestCase("Data/multipleemptysheets.xlsx")]
    public void A010_EmptyXlsx(string fileName)
    {
        using IExcelFast workbook = new ExcelFastXlsx();
        workbook.Open(fileName);
        workbook.SheetNames().Should().NotBeEmpty();
    }

    [Test]
    [TestCase("Data/invalidfile.xlsx")]
    public void A020_NonZipFile(string fileName)
    {
        Action sutMethod =  () =>
        {
            using IExcelFast workbook = new ExcelFastXlsx();
            workbook.Open(fileName);
        };

        sutMethod.Should().Throw<InvalidDataException>();
    }

    [Test]
    [TestCase("Data/missingworkbook.xlsx")]
    //[TestCase("Data/missingworkbookrelatioship.xlsx")] This can be loaded by "LibreOffice Calc" !!
    public void A030_InvalidXlsx(string fileName)
    {
        Action sutMethod =  () =>
        {
            using IExcelFast workbook = new ExcelFastXlsx();
            workbook.Open(fileName);
        };
        sutMethod.Should().Throw<ArgumentNullException> ();
    }

    [Test]
    [TestCase("Data/verysimple.xlsx")]
    public void A040_DisposeReleasesFile(string fileName)
    {
        using (IExcelFast workbook = new ExcelFastXlsx())
        {
            workbook.Open(fileName);
            //read lock is held
            Assert.Throws<IOException>(() => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
        }
        //no lock, open read/write should work
        using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

}

