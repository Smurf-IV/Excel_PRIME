using System;
using System.Runtime.CompilerServices;
using System.Xml;

namespace ExcelPRIME.Implementation;

// ReSharper disable InconsistentNaming on "private const string"s
internal sealed class SharedStringsRestrictedNameTable : NameTable
{
    private const string tAtom = "t";
    private const string siAtom = "si";
    private const string sstAtom = "sst";
    private const string uniqueCountAtom = "uniqueCount";

    public SharedStringsRestrictedNameTable()
    {
        // TODO: Check if adding these on construction gives a benefit.
        //["sst", "si", "t", "r", "rPr", "b", "sz", "mc", "color", "rFont", "family", "charset",
        //"xml", "xmlns", string.Empty, "http://www.w3.org/2000/xmlns/", "http://www.w3.org/XML/1998/namespace", "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        //"version", "encoding", "standalone", "count", "uniqueCount", "val", "rgb", "space", "i"]
    }

    public override string Add(char[] key, int start, int len) => Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);

    public override string Add(string key) => Get(key.AsSpan()) ?? base.Add(key);

    public override string? Get(char[] key, int start, int len) => Get(key.AsSpan(start, len));

    public override string? Get(string value) => Get(value.AsSpan());

    [MethodImpl(MethodImplOptions.AggressiveOptimization | MethodImplOptions.AggressiveInlining)]
    private static string? Get(ReadOnlySpan<char> value)
    {
        switch (value.Length)
        {
            case 0:
                return string.Empty;
            case 1:
                if (value[0] == tAtom[0]) return tAtom;
                break;
            case 2:
                if (value.SequenceEqual(siAtom)) return siAtom;
                break;
            case 3:
                if (value.SequenceEqual(sstAtom)) return sstAtom;
                break;
            case 11:
                if (value.SequenceEqual(uniqueCountAtom)) return uniqueCountAtom;
                break;
        }
        return null;
    }
}

internal sealed class SheetRestrictedNameTable : NameTable
{
    private const string cAtom = "c";
    private const string rAtom = "r";
    private const string tAtom = "t";
    private const string sAtom = "s";
    private const string vAtom = "v";
    private const string isAtom = "is";
    private const string rowAtom = "row";
    private const string colAtom = "col";
    private const string refAtom = "ref";
    private const string hiddenAtom = "hidden";
    private const string dimensionAtom = "dimension";
    private const string sheetDataAtom = "sheetData";
    private const string worksheetAtom = "worksheet";

    public override string Add(char[] key, int start, int len) => Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);

    public override string Add(string key) => Get(key.AsSpan()) ?? base.Add(key);

    public override string? Get(char[] key, int start, int len) => Get(key.AsSpan(start, len));

    public override string? Get(string value) => Get(value.AsSpan());

    [MethodImpl(MethodImplOptions.AggressiveOptimization | MethodImplOptions.AggressiveInlining)]
    private static string? Get(ReadOnlySpan<char> value)
    {
        switch (value.Length)
        {
            case 0:
                return string.Empty;
            case 1:
                switch (value[0])
                {
                    case 'c': return cAtom;
                    case 'r': return rAtom;
                    case 't': return tAtom;
                    case 's': return sAtom;
                    case 'v': return vAtom;
                }
                break;
            case 2:
                if (value.SequenceEqual(isAtom)) return isAtom;
                break;
            case 3:
                // ReSharper disable once ConvertIfStatementToSwitchStatement
                if (value.SequenceEqual(rowAtom)) return rowAtom;
                if (value.SequenceEqual(colAtom)) return colAtom;
                if (value.SequenceEqual(refAtom)) return refAtom;
                break;
            case 6:
                if (value.SequenceEqual(hiddenAtom)) return hiddenAtom;
                break;
            case 9:
                // ReSharper disable once ConvertIfStatementToSwitchStatement
                if (value.SequenceEqual(dimensionAtom)) return dimensionAtom;
                if (value.SequenceEqual(sheetDataAtom)) return sheetDataAtom;
                if (value.SequenceEqual(worksheetAtom)) return worksheetAtom;
                break;
        }
        return null;
    }
}
