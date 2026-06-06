using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
[TestFixture]
internal class LoadingTestsXlsb
{
    [Test]
    [TestCase("test_not_exist.xlsb")]
    public async Task A000_FileNotExist_ThrowsFileNotFoundException(string fileName)
    {
        Func<Task> sutMethod = async () =>
        {
            using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
            await workbook.OpenAsync(fileName).ConfigureAwait(false);
        };

        await sutMethod.Should().ThrowAsync<FileNotFoundException>()
            .WithMessage("Could not find file *").ConfigureAwait(false);
    }

    [Test]
    [TestCase("Data/empty.xlsb")]
    [TestCase("Data/multipleemptysheets.xlsb")]
    public async Task A010_EmptyXlsx(string fileName)
    {
        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
        await workbook.OpenAsync(fileName).ConfigureAwait(false);
        workbook.SheetNames().Should().NotBeEmpty();
    }

    [Test]
    [TestCase("Data/invalidfile.xlsx")]
    public async Task A020_NonZipFile(string fileName)
    {
        Func<Task> sutMethod = async () =>
        {
            using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
            await workbook.OpenAsync(fileName).ConfigureAwait(false);
        };

        await sutMethod.Should().ThrowAsync<InvalidDataException>().ConfigureAwait(false);
    }

    //[Test]
    //[TestCase("Data/missingworkbook.xlsb")]
    ////[TestCase("Data/missingworkbookrelatioship.xlsb")] This can be loaded by "LibreOffice Calc" !!
    //public async Task A030_InvalidXlsx(string fileName)
    //{
    //    Func<Task> sutMethod = async () =>
    //    {
    //        using IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb();
    //        await workbook.OpenAsync(fileName).ConfigureAwait(false);
    //    };
    //    await sutMethod.Should().ThrowAsync<ArgumentNullException> ().ConfigureAwait(false);
    //}

    [Test]
    [TestCase("Data/verysimple.xlsb")]
    public async Task A040_DisposeReleasesFile(string fileName)
    {
        using (IExcel_PRIMEAsync workbook = new Excel_PRIMEXlsb())
        {
            await workbook.OpenAsync(fileName).ConfigureAwait(false);
            //read lock is held
            Assert.Throws<IOException>(() => File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
        }
        //no lock, open read/write should work
        using FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
    }

}

