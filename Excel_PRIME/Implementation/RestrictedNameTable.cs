using System;
using System.Buffers;
using System.Linq;
using System.Xml;

namespace ExcelPRIME.Implementation;

sealed class SharedStringsRestrictedNameTable : NameTable
{
    private const string tAtom = "t";
    private const string siAtom = "si";
    private const string sstAtom = "sst";
    private const string uniqueCountAtom = "uniqueCount";

    public SharedStringsRestrictedNameTable()
    {
        //["sst", "si", "t", "r", "rPr", "b", "sz", "mc", "color", "rFont", "family", "charset",
        //"xml", "xmlns", string.Empty, "http://www.w3.org/2000/xmlns/", "http://www.w3.org/XML/1998/namespace", "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        //"version", "encoding", "standalone", "count", "uniqueCount", "val", "rgb", "space", "i"]
    }

    public override string Add(char[] key, int start, int len) => Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);

    public override string Add(string key) => Get(key.AsSpan()) ?? base.Add(key);

    public override string? Get(char[] key, int start, int len) => Get(key.AsSpan(start, len));

    public override string? Get(string value) => Get(value.AsSpan());

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

sealed class SheetRestrictedNameTable : NameTable
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
    //private const string minAtom = "min";
    //private const string maxAtom = "max";
    private const string hiddenAtom = "hidden";
    private const string dimensionAtom = "dimension";
    private const string sheetDataAtom = "sheetData";
    private const string worksheetAtom = "worksheet";

    public override string Add(char[] key, int start, int len) => Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);

    public override string Add(string key) => Get(key.AsSpan()) ?? base.Add(key);

    public override string? Get(char[] key, int start, int len) => Get(key.AsSpan(start, len));

    public override string? Get(string value) => Get(value.AsSpan());

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
                switch (value)
                {
                    case rowAtom:
                        return rowAtom;
                    case colAtom:
                        return colAtom;
                    case refAtom:
                        return refAtom;
                }

                //if (value.SequenceEqual(minAtom)) return minAtom;
                //if (value.SequenceEqual(maxAtom)) return maxAtom;
                break;
            //case 4:
            //    if (value.SequenceEqual("cols")) return "cols";
            //    break;
            //case 5:
            //    if (value.SequenceEqual("spans")) return "spans";
            //    break;
            case 6:
                if (value.SequenceEqual(hiddenAtom)) return hiddenAtom;
                break;
            case 9:
                switch (value)
                {
                    case dimensionAtom:
                        return dimensionAtom;
                    case sheetDataAtom:
                        return sheetDataAtom;
                    case worksheetAtom:
                        return worksheetAtom;
                }

                //if (value.SequenceEqual("dyDescent")) return "dyDescent";
                break;
        }
        return null;
    }
}
