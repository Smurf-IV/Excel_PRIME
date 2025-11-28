using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;

using BenchmarkDotNet.Attributes;

using Microsoft.VisualBasic.CompilerServices;

namespace ExcelPRIME.RangeBench;

[ExcludeFromCodeCoverage]
public class RangeBenchmarks
{
    public IEnumerable<Type> Rangers() // for multiple arguments it's an IEnumerable of IGetRange's
    {
        yield return typeof(GetRangeExcelPrime);
        yield return typeof(GetRangeClosedXML);
        yield return typeof(GetRangeEPPlus);
        yield return typeof(GetRangeFreeSpire);
        yield return typeof(GetRangeAsposeCells);
    }


    [Benchmark]
    [ArgumentsSource(nameof(Rangers))]
    [MethodImpl(MethodImplOptions.NoOptimization)]
    public int Access100mb(Type ranger)
    {
        using IGetRange getRanger = (IGetRange)Activator.CreateInstance(ranger)!;
        getRanger.LoadFile("Data\\100mb.xlsx");

        //< definedName name = "DışVeri_1" localSheetId = "2" hidden = "1" > Tablo3!$A$1:$H$99929 </ definedName >
        //< definedName name = "DışVeri_1" localSheetId = "3" hidden = "1" > Worksheet!$A$938995:$H$952350 </ definedName >
        //< definedName name = "DışVeri_1" localSheetId = "0" hidden = "1" > 'Worksheet (2)'!$A$1:$H$27001 </ definedName >
        //< definedName name = "DışVeri_1" localSheetId = "1" hidden = "1" > 'Worksheet (3)'!$A$1:$H$4001 </ definedName >
        //< definedName name = "DışVeri_2" localSheetId = "3" hidden = "1" > Worksheet!$A$952351:$H$985351 </ definedName >
        var rangeTablo3 = getRanger.GetDefinedRange("DışVeri_1", 2);
        int cells = rangeTablo3.Sum(row => row.Count());
        var rangeWorksheet = getRanger.GetDefinedRange("DışVeri_1", 3);
        cells += rangeWorksheet.Sum(row => row.Count());
        var rangeWorksheet2 = getRanger.GetDefinedRange("DışVeri_1", 0);
        cells += rangeWorksheet2.Sum(row => row.Count());
        var rangeWorksheet3 = getRanger.GetDefinedRange("DışVeri_1", 1);
        cells += rangeWorksheet3.Sum(row => row.Count());
        var rangeDışVeri_2 = getRanger.GetDefinedRange("DışVeri_2", 3);
        cells += rangeDışVeri_2.Sum(row => row.Count());

        return cells;
    }
}
