using System;
using System.Diagnostics.CodeAnalysis;
using System.Threading.Tasks;

using JetBrains.dotMemoryUnit;

using NUnit.Framework;


namespace ExcelPRIME.Tests;

[ExcludeFromCodeCoverage]
internal class DotMemoryUnit : IDisposable
{
    public static IDisposable Support => new DotMemoryUnit();

    private DotMemoryUnit()
    {
        DotMemoryUnitController.TestStart();
    }

    public void Dispose() => DotMemoryUnitController.TestEnd();
}

[ExcludeFromCodeCoverage]
[NonParallelizable]
internal class BugTesting
{
    [Test]
    [Explicit("Run DotMemory")]
    [DotMemoryUnit(CollectAllocations = true, FailIfRunWithoutSupport=true, Directory = @".\DotMemory")]

    public void Bug_000_SharedStrings_DotMemory()
    {
        using IDisposable dms = DotMemoryUnit.Support;
        Bug_001_SharedStrings().GetAwaiter().GetResult();
    }

    [Test]
    public async Task Bug_001_SharedStrings()
    {
        const string fileName = "Data/100mb.xlsx";
        int cells = 0;
        using IExcel_PRIME workbook = new Excel_PRIME();
        await workbook.OpenAsync(fileName).ConfigureAwait(true);
        foreach (string sheetName in workbook.SheetNames())
        {
            using ISheet? worksheet = await workbook.GetSheetAsync(sheetName).ConfigureAwait(false);
            foreach (IRow? row in worksheet!.GetRowData())
            {
                if (row == null)
                {   // Because this returns upto the dimension of the sheet Height
                    break;
                }

                foreach (ICell? cell in row.GetAllCells())
                {
                    cells++;
                    if (cells >= 216017)
                    {
                        // When requesting "922473" it should not be returning a blank string!
                        //cell.RawValue.ToString().Should().NotBeNullOrWhiteSpace();
                    }
                }
                row.Dispose();
            }
        }

    }
}
