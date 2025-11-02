using System.Xml;
// ReSharper disable InconsistentNaming

namespace ExcelPRIME.Implementation;

internal record struct ReaderAtoms
{
    // Row
    public string rowRefAtom { get; }
    public string hiddenRefAtom { get; }
    public string cRefAtom { get; }
    public string rRefAtom { get; }
    public string tRefAtom { get; }
    public string vRefAtom { get; }
    public string sRefAtom { get; }
    
    public ReaderAtoms(XmlReader reader)
    {
        // Row
        rowRefAtom = reader.NameTable.Add("row");
        rRefAtom = reader.NameTable.Add("r");
        hiddenRefAtom = reader.NameTable.Add("hidden");
        cRefAtom = reader.NameTable.Add("c");
        // Cell
        //rRefAtom = reader.NameTable.Add("r");
        tRefAtom = reader.NameTable.Add("t");
        vRefAtom = reader.NameTable.Add("v");
        sRefAtom = reader.NameTable.Add("s");
    }
}