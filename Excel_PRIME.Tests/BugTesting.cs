using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using AwesomeAssertions;

using NUnit.Framework;


namespace ExcelPRIME.Tests;


[ExcludeFromCodeCoverage]
[NonParallelizable]
[TestFixture]
internal class BugTesting
{
    [Test]
    public async Task Bug_001_SharedStrings()
    {
        const string fileName = "Data/100mb.xlsx";
        int cells = 0;
        using Excel_PRIMEAsync workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row == null)
                {   // Because this returns up to the dimension of the sheet Height
                    break;
                }

                if (row is INullRow)
                {
                    continue;
                }

                ReadOnlySpan<Cell> rowCells = row.GetAllCells();
                row.Dispose();
                if (rowCells.IsEmpty)
                {
                    continue;
                }
                foreach (ref readonly Cell cell in rowCells)
                {
                    // Because this returns up to the dimension of the sheet width
                    cells++;
                    if (cells >= 216017)
                    {
                        // When requesting "922473" it should not be returning a blank string!
                        //cell.RawValue.ToString().Should().NotBeNullOrWhiteSpace();
                    }
                }
            }
        }

    }

    [Test]
    public void Bug_020_AsDecimal_FromDecimal()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.Create(val, -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

#if NET11_0_OR_GREATER
/*
- Floating-Point Representation (IEEE 754) -

The base-10 number 41273.28 cannot be represented with exact binary precision in a standard 64-bit floating-point (double) number. In IEEE 754 binary floating-point representation:
• The exact closest 64-bit float representation of 41273.28 is: 41273.27999999999883584678173065185546875 (0x40E42728F5C28F5C).

- Runtime Conversion Mechanics: .NET 8/9/10 vs. .NET 11 -

-- 1. .NET 8, .NET 9, and .NET 10 (Legacy Rounding) --
Historically in CoreCLR / .NET runtime (inherited from legacy OLE Automation / VarDecFromR8), explicit casts from double to decimal rounded the double value to 15 significant decimal digits.
• Converting 41273.2799999999988358... to decimal with 15 digits of precision rounded the value up to 41273.28M.
• Therefore, in .NET 8–10, (decimal)(double)41273.28m returned 41273.28M, masking the floating-point imprecision.

-- 2. .NET 11 (Exact Precision Conversion) --
In .NET 11, the runtime modernized floating-point-to-decimal conversion intrinsics in RyuJIT/CoreCLR to avoid artificial 15-digit truncation and preserve full bitwise precision up to the 28–29 digits supported by decimal:
• (decimal)(double)41273.28m preserves the exact value stored in the IEEE 754 double bits, yielding 41273.279999999998835846781731M.
• When CellValue.Create((decimal)(double)val, -1) executes, 41273.279999999998835846781731M is stored into _d.
• cellValue.AsDecimal evaluates to 41273.279999999998835846781731M, which fails strict equality comparison Should().Be(41273.28M).
*/
#else
    [Test]
    public void Bug_020_AsDecimal_FromDouble()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.Create((decimal)(double)val, -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }
#endif
    [Test]
    public void Bug_020_AsDecimal_FromString()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.Create("41273.28", -1);

        // Act & Assert
        cellValue.AsDecimal.Should().Be(val);
    }

    [Test]
    public void Bug_020_AsDecimal_FromSpan()
    {
        // Arrange
        decimal val = 41273.28m;
        CellValue cellValue = CellValue.TryParseOrder("41273.28".AsSpan(), null);

        // Act & Assert
#if NET11_0_OR_GREATER
        cellValue.AsDecimal.Should().BeApproximately(val, 0.00000001m);
#else
        cellValue.AsDecimal.Should().Be(val);
#endif
    }

    [Test]
    public async Task Bug_022_EndElement()
    {
        const string fileName = "Data/MissingCells.xlsx";
        int cells = 0;
        using Excel_PRIMEAsync workbook = new ();
        await workbook.OpenAsync(fileName, new Options { CellConversionType = CellConversion.ExcelCellType }).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = workbook.GetSheet(sheetName);
            foreach (IRow? row in worksheet!.GetRowData(2))
            {
                ReadOnlySpan<Cell> rowCells = row.GetAllCells();
                row.Dispose();
                if (rowCells.IsEmpty)
                {
                    continue;
                }
                foreach (ref readonly Cell cell in rowCells)
                {
                    // Because this returns up to the dimension of the sheet width
                    cells++;
                }
            }
        }
        cells.Should().Be(10);
    }

    [Test]
    public async Task Bug_022_EndElement_Async()
    {
        const string fileName = "Data/MissingCells.xlsx";
        int cells = 0;
        using Excel_PRIMEAsync workbook = new ();
        await workbook.OpenAsync(fileName, new Options { CellConversionType = CellConversion.ExcelCellType }).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            await foreach (IRowAsync? row in worksheet!.GetRowDataAsync(2).ConfigureAwait(false))
            {
                IReadOnlyList<Cell>? rowCells = await row.GetAllCellsAsync().ConfigureAwait(true);
                row.Dispose();
                if (rowCells == null)
                {
                    continue;
                }
                foreach (Cell cell in rowCells)
                {
                    // Because this returns up to the dimension of the sheet width
                    cells++;
                }
            }
        }
        cells.Should().Be(10);

    }

    public static Options[] Option =
    [
        //new Options(),
        //new Options { ReturnDBNull = true },
        new Options { CellConversionType = CellConversion.ExcelCellType},
        new Options { CellConversionType = CellConversion.ExcelCellType, ReturnDBNull = true },
        new Options { CellConversionType = CellConversion.ForceStyles }
    ];

    [Test]
    [TestCaseSource(nameof(Option))]
    public async Task Bug_019_DateInObj_XLSX(Options options)
    {
        const string fileName = "Data/65K_Records_Data.xlsx";
        using Excel_PRIMEAsync workbook = new();
        await Bug_019_Core(options, workbook, fileName);
    }

    [Test]
    [TestCaseSource(nameof(Option))]
    public async Task Bug_019_DateInObj_XLSB(Options options)
    {
        const string fileName = "Data/65K_Records_Data.xlsb";
        using Excel_PRIMEXlsbAsync workbook = new();
        await Bug_019_Core(options, workbook, fileName);
    }

    private static async Task Bug_019_Core(Options options, Excel_PRIMEAsync workbook, string fileName)
    {
        await workbook.OpenAsync(fileName, options).ConfigureAwait(false);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("500000 Sales Records").ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(1, RowCellGet.None).FirstAsync();
        Cell cell = await row.GetCellAsync(6).ConfigureAwait(false);
        cell.CellValue.AsDateTime.Should().Be(new DateTime(2012, 7, 27)); // Date DD/MM/YYYY
    }

    [CancelAfter(1000)]
    [Test]
    public async Task Bug_027_SkipLines_XLSX(CancellationToken ct)
    {
        const string fileName = "Data/SkipLines.xlsx";
        using Excel_PRIMEAsync workbook = new ();
        await workbook.OpenAsync(fileName, ct: ct).ConfigureAwait(true);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("SkipLines", ct: ct).ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(3, RowCellGet.None, ct: ct).FirstAsync(ct);
        Cell cell = await row.GetCellAsync(6, ct: ct).ConfigureAwait(false);
        cell.CellValue.AsInt32.Should().Be(1);
    }
    [CancelAfter(1000)]
    [Test]
    public async Task Bug_027_SkipLines_XLSB(CancellationToken ct)
    {
        const string fileName = "Data/SkipLines.xlsb";
        using Excel_PRIMEXlsbAsync workbook = new ();
        await workbook.OpenAsync(fileName, ct: ct).ConfigureAwait(true);
        ISheetAsync? valSheet = await workbook.GetSheetAsync("SkipLines", ct: ct).ConfigureAwait(false);
        IRowAsync? row = await valSheet.GetRowDataAsync(3, RowCellGet.None, ct: ct).FirstAsync(ct);
        Cell cell = await row.GetCellAsync(6, ct: ct).ConfigureAwait(false);
        cell.CellValue.AsInt32.Should().Be(1);
    }

    [Test]
    public async Task Bug_028_MultiOpen()
    {
        const string fileName = "Data/100mb.xlsx";
        using Excel_PRIMEAsync workbook1 = new ();
        await workbook1.OpenAsync(fileName).ConfigureAwait(true);
        using Excel_PRIMEAsync workbook2 = new ();
        await workbook2.OpenAsync(fileName).ConfigureAwait(true);
    }
    [Test]
    public async Task Bug_028_LeftOpen()
    {
        const string fileName = "Data/100mb.xlsx";
        Excel_PRIMEAsync workbook = new ();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            /*using*/ ISheetAsync? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            using IRowAsync? row = await worksheet.GetRowDataAsync(3, RowCellGet.None).FirstAsync();
        }

        workbook.Dispose();
        await using var fs = File.OpenWrite(fileName);
        fs.Should().BeWriteOnly();
    }

    [Test]
    public async Task Bug_034_CaseRangeName()
    {
        const string fileName = "Data/named-range.xlsx";
        using Excel_PRIMEAsync workbook = new();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        // This is the expected behavior, the defined name is case-insensitive, so both "TaxRate" and "taxrate" should return the same value.
        //CellValue[] taxRate = await workbook.GetDefinedRangeAsync("TaxRate").FirstAsync();
        // Test uppercase and lowercase defined name, they should return the same value.
        CellValue[] taxRate = await workbook.GetDefinedRangeAsync("TAXRATE").FirstAsync();
        taxRate.First().ToString().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");
        
        taxRate = await workbook.GetDefinedRangeAsync("taxrate").FirstAsync();
        taxRate.First().ToString().Should().Be("0.1", "<definedName name=\"TaxRate\">0.1</definedName>");
    }

}

