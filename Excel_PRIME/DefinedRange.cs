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
    internal DefinedRange(string reference)
    {
        Name = string.Empty;
        ReadOnlySpan<char> span = reference.AsSpan();
        int exclIndex = span.IndexOf('!');
        if (exclIndex >= 0)
        {
            // sheet name may be quoted with '
            ReadOnlySpan<char> sheetSpan = span.Slice(0, exclIndex);
            if (sheetSpan.Length > 0 && sheetSpan[0] == '\'')
            {
                // Trim leading and trailing single quote
                if (sheetSpan.Length >= 2 && sheetSpan[^1] == '\'')
                {
                    SheetName = new string(sheetSpan.Slice(1, sheetSpan.Length - 2));
                }
                else
                {
                    SheetName = sheetSpan.ToString();
                }
            }
            else
            {
                SheetName = sheetSpan.ToString();
            }

            ReadOnlySpan<char> range = span.Slice(exclIndex + 1);
            if (range.IndexOf(':') >= 0)
            {
                DoExtractBasedOnCellRange(range);
            }
            else if (range.IndexOf('$') >= 0)
            {
                DoExtractBasedOnSingleCell(range);
            }
        }
        else if (span.IndexOf(':') >= 0)
        {
            DoExtractBasedOnCellRange(span);
        }
        else if (span.IndexOf('$') >= 0)
        {
            DoExtractBasedOnSingleCell(span);
        }
        else
        {
            // e.g <definedName name="TaxRate">0.1</definedName>
            ConstValue = reference;
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
    public DefinedRange(string sheetName, ReadOnlySpan<char> columnStart, ReadOnlySpan<char> columnEnd, int rowStart = 1, int rowEnd = 1_048_576)
    {
        SheetName = sheetName;
        Name = string.Empty;
        ExcelColumnStart = columnStart.GetColNumber();
        ExcelColumnEnd = columnEnd.GetColNumber();
        ExcelRowStart = rowStart;
        ExcelRowEnd = rowEnd;
    }

    /// <summary>
    /// User defined range (With or Without `$`'s, e.g., `A1:B2`)
    /// </summary>
    /// <param name="userRange"></param>
    /// <param name="sheetName"></param>
    public DefinedRange(string userRange, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(userRange))
        {
            throw new ArgumentException("userRange cannot be null or whitespace", nameof(userRange));
        }
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException("sheetName cannot be null or whitespace", nameof(sheetName));
        }

        SheetName = sheetName;
        Name = userRange;
        if (userRange.Contains(':'))
        {
            DoExtractBasedOnCellRange(userRange);
        }
        else 
        {
            DoExtractBasedOnSingleCell(userRange);
        }
    }
    
    private void DoExtractBasedOnSingleCell(ReadOnlySpan<char> range) // e.g. "$C$12" or span of that
    {
        // Expect patterns like $C$12 or maybe C12
        // Find first '$' then next '$'
        int firstDollar = range.IndexOf('$');
        int secondDollar = -1;
        if (firstDollar >= 0)
        {
            secondDollar = range.Slice(firstDollar + 1).IndexOf('$');
            if (secondDollar >= 0)
            {
                secondDollar += firstDollar + 1;
            }
        }

        if (firstDollar >= 0 && secondDollar > firstDollar)
        {
            ReadOnlySpan<char> colSpan = range.Slice(firstDollar + 1, secondDollar - (firstDollar + 1));
            ReadOnlySpan<char> rowSpan = range.Slice(secondDollar + 1);
            ExcelColumnStart = ExcelColumnEnd = colSpan.GetColNumber();
            ExcelRowStart = ExcelRowEnd = rowSpan.IntParse();
        }
        else
        {
            // fallback: try parsing letters then digits
            int i = 0;
            while (i < range.Length && !char.IsDigit(range[i])) i++;
            ReadOnlySpan<char> col = range.Slice(0, i);
            ReadOnlySpan<char> row = range.Slice(i);
            ExcelColumnStart = ExcelColumnEnd = col.GetColNumber();
            ExcelRowStart = ExcelRowEnd = row.IntParse();
        }
    }

    private void DoExtractBasedOnCellRange(ReadOnlySpan<char> range) // e.g. "$C$12:$E$12" or similar
    {
        // Default value
        ExcelRowStart = 1;
        int colon = range.IndexOf(':');
        ReadOnlySpan<char> left = colon >= 0 ? range.Slice(0, colon) : range;
        ReadOnlySpan<char> right = colon >= 0 ? range.Slice(colon + 1) : [];

        // parse left
        if (left.Length > 0)
        {
            int firstDollar = left.IndexOf('$');
            int secondDollar = -1;
            if (firstDollar >= 0)
            {
                secondDollar = left.Slice(firstDollar + 1).IndexOf('$');
                if (secondDollar >= 0)
                {
                    secondDollar += firstDollar + 1;
                }
            }

            if (firstDollar >= 0 && secondDollar > firstDollar)
            {
                ReadOnlySpan<char> startCol = left.Slice(firstDollar + 1, secondDollar - (firstDollar + 1));
                ExcelColumnStart = startCol.GetColNumber();
                if (secondDollar + 1 < left.Length)
                {
                    ExcelRowStart = left.Slice(secondDollar + 1).IntParse();
                }
            }
            else
            {
                int i = 0; while (i < left.Length && !char.IsDigit(left[i])) i++;
                ExcelColumnStart = left.Slice(0, i).GetColNumber();
                if (i < left.Length)
                {
                    ExcelRowStart = left.Slice(i).IntParse();
                }
            }
        }

        // parse right
        if (right.Length > 0)
        {
            int firstDollar = right.IndexOf('$');
            int secondDollar = -1;
            if (firstDollar >= 0)
            {
                secondDollar = right.Slice(firstDollar + 1).IndexOf('$');
                if (secondDollar >= 0)
                {
                    secondDollar += firstDollar + 1;
                }
            }

            if (firstDollar >= 0 && secondDollar > firstDollar)
            {
                ReadOnlySpan<char> endCol = right.Slice(firstDollar + 1, secondDollar - (firstDollar + 1));
                ExcelColumnEnd = endCol.GetColNumber();
                if (secondDollar + 1 < right.Length)
                {
                    ExcelRowEnd = right.Slice(secondDollar + 1).IntParse();
                }
            }
            else
            {
                int i = 0; while (i < right.Length && !char.IsDigit(right[i])) i++;
                ExcelColumnEnd = right.Slice(0, i).GetColNumber();
                if (i < right.Length)
                {
                    ExcelRowEnd = right.Slice(i).IntParse();
                }
            }
        }
    }

    /// <summary>
    /// Xml.Attribute("name"); 
    /// </summary>
    public string Name { get; init; }

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

    ///// <summary>
    ///// Xml;
    ///// </summary>
    //public string SheetIdReference { get; init; }

    /// <summary>
    /// Xml;
    /// </summary>
    public string? ConstValue { get; private set; }
}
