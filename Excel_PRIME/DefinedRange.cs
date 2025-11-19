using System;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME;

/// <summary>
/// Defines a range of cells using a reference string
/// </summary>
public record DefinedRange
{
    /// <summary>
    /// Defines a range of cells using a reference string
    /// </summary>
    /// <param name="reference">Reference string i.e. Sheet1!$A$1:$A$4</param>
    /// <exception cref="ArgumentException">Thrown when reference is invalid or not supported</exception>
    internal DefinedRange(in string reference)
    {
        if (reference.Contains('('))
        {
            // e.g <definedName name="Prices">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>
            ConstValue = reference;
        }
        else if (reference.Contains('!'))
        {
            DoExtractBasedOnSheetName(reference);
        }
        else if (reference.Contains(':'))
        {
            // e.g $A$1:$A$4
            DoExtractBasedOnCellRange(reference);
        }
        else if (reference.Contains('$'))
        {
            // e.g. $A$1
            DoExtractBasedOnSingleCell(reference);
        }
        else
        {
            // e.g <definedName name="TaxRate">0.1</definedName>
            ConstValue = reference;
        }
    }

    private void DoExtractBasedOnSheetName(in string reference)
    {
        string[] splitReference = reference.Split('!');

        SheetName = splitReference[0].Trim('\'');
        string range = splitReference[1];

        if (range.Contains(':'))     // "$C$12:$E$12"
        {
            DoExtractBasedOnCellRange(range);
        }
        else if (range.Contains('$'))   // "$C$12"
        {
            DoExtractBasedOnSingleCell(range);
        }
    }

    private void DoExtractBasedOnSingleCell(in string range) // "$C$12"
    {
        string[] strings = range.Split('$');
        ExcelColumnStart = ExcelColumnEnd = strings[1].GetColNumber();
        ExcelRowEnd = ExcelRowStart = strings[2].IntParse();
    }

    private void DoExtractBasedOnCellRange(in string range) // "$C$12:$E$12"
    {
        // Default value
        ExcelRowStart = 1;
        string[] splitRange = range.Split(':');
        //if (splitRange.Length > 1)
        {
            string[] startRef = splitRange[0].Split('$');
            ExcelColumnStart = startRef[1].GetColNumber();  // Start with a '$', therefore first entry is empty
            string[] endRef = splitRange[1].Split('$');
            ExcelColumnEnd = endRef[1].GetColNumber();
            if (startRef.Length > 2)
            {
                ExcelRowStart = startRef[2].IntParse();
            }
            if (endRef.Length > 2)
            {
                ExcelRowEnd = endRef[2].IntParse();
            }
        }
    }

    /// <summary>
    /// Defines a cell range using variables
    /// </summary>
    /// <param name="columnStart">Column Letter start</param>
    /// <param name="columnEnd">Column Letter end</param>
    /// <param name="rowStart">First row number</param>
    /// <param name="rowEnd">last row number [Excel 2010 specifies 1_048_576, but Power Query can go upto  1_999_999_997]</param>
    /// <param name="sheetName">The Sheet this will be applied to.</param>
    public DefinedRange(in string sheetName, ReadOnlySpan<char> columnStart, ReadOnlySpan<char> columnEnd, int rowStart = 1, int rowEnd = 1_048_576)
    {
        SheetName = sheetName;
        ExcelColumnStart = columnStart.GetColNumber();
        ExcelColumnEnd = columnEnd.GetColNumber();
        ExcelRowStart = rowStart;
        ExcelRowEnd = rowEnd;
    }

    /// <summary>
    /// Xml.Attribute("name").Value; 
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    /// What SheetName (If Any) does this belong to
    /// </summary>
    public string? SheetName { get; private set; }

    /// <summary>
    /// Column Range Start
    /// </summary>
    public int ExcelColumnStart { get; private set; }

    /// <summary>
    /// Column Range End
    /// </summary>
    public int ExcelColumnEnd { get; private set; }

    /// <summary>
    /// Row Range Start
    /// </summary>
    public int ExcelRowStart { get; private set; }

    /// <summary>
    /// Row Range End
    /// </summary>
    public int ExcelRowEnd { get; private set; }

    /// <summary>
    /// Xml.Value;
    /// </summary>
    public required string SheetIdReference { get; init; }

    /// <summary>
    /// Xml.Value;
    /// </summary>
    public string? ConstValue { get; private set; }
}
