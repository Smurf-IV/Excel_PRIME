using System;
using System.Linq;

using ExcelPRIME.Shared;

namespace ExcelPRIME;

public record DefinedRange
{
    /// <summary>
    /// Defines a range of cells using a reference string
    /// </summary>
    /// <param name="reference">Reference string i.e. Sheet1!$A$1</param>
    /// <exception cref="ArgumentException">Thrown when reference is invalid or not supported</exception>
    internal DefinedRange(string reference)
    {
        string[] splitReference = reference.Split('!');

        SheetName = splitReference[0];
        string range = splitReference[1];

        // Default value
        RowStart = 1;

        string[] splitRange = range.Split(':');
        if (splitRange.Length > 1)
        {
            string[] startRef = splitRange[0].Split('$');
            ColumnStart = startRef[1].IntParseUnsafe();
            string[] endRef = splitRange[1].Split('$');
            ColumnEnd = endRef[1].IntParseUnsafe();
            if (splitRange[0].Count(c => c == '$') > 1)
            {
                RowStart = startRef[2].IntParseUnsafe();
                RowEnd = endRef[2].IntParseUnsafe();
            }
        }
        else
        {
            ColumnStart = ColumnEnd = range.Split('$')[1].IntParseUnsafe();
            RowEnd = RowStart = range.Split('$')[2].IntParseUnsafe();
        }
    }

    /// <summary>
    /// Defines a cell range using variables
    /// </summary>
    /// <param name="columnStart">Column Letter start</param>
    /// <param name="columnEnd">Column Letter end</param>
    /// <param name="rowStart">First row number</param>
    /// <param name="rowEnd">last row number</param>
    public DefinedRange(string sheetName, int columnStart, int columnEnd, int rowStart = 1, int? rowEnd = null)
    {
        SheetName = sheetName;
        ColumnStart = columnStart;
        ColumnEnd = columnEnd;
        RowStart = rowStart;
        RowEnd = rowEnd;
    }
    /// <summary>
    /// Xml.Attribute("name").Value; 
    /// </summary>
    public required string Name { get; init; }

    private readonly string? SheetName;

    /// <summary>
    /// Column Range Start
    /// </summary>
    public int? ColumnStart { get; set; }
    /// <summary>
    /// Column Range End
    /// </summary>
    public int? ColumnEnd { get; set; }
    /// <summary>
    /// Row Range Start
    /// </summary>
    public int? RowStart { get; set; }
    /// <summary>
    /// Row Range End
    /// </summary>
    public int? RowEnd { get; set; }
    /// <summary>
    /// Xml.Value;
    /// </summary>
    public required string Reference { get; init; }

    /// <summary>
    /// Used to generate a key for the dictionary 
    /// </summary>
    public string Key => Name + SheetName;
}
