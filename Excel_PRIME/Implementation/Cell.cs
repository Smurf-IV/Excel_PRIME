using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class Cell : ICell
{
    private bool _isDisposed;

    public Cell(XmlReader reader, ISharedString sharedStrings)
    {
        string address = string.Empty;// = reader.GetAttribute("r");
        string type = string.Empty;// = reader.GetAttribute("t");
        int needed = 0;
        while (reader.MoveToNextAttribute())
        {
            switch (reader.LocalName)
            {
                case "r":
                    address = reader.Value;
                    needed++;
                    break;
                case "t":
                    type = reader.Value;
                    needed++;
                    break;
            }
        }
        // var style = reader.GetAttribute("s"); Use this when formatting the value
        bool isTextRow = type == "s";

        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        (int _, int col, ReadOnlyMemory<char> colName) = address.GetRowColNumbers();
        ColumnLetters = colName;
        //RowNumber = row;
        ExcelColumnOffset = col;
        string? value = null;
        if (reader.Read() 
            && reader is { IsEmptyElement: false, LocalName: "v" })
        {
            if (reader.IsStartElement() 
                || reader.MoveToContent() == XmlNodeType.Element
                )
            {
                value = reader.ReadString();
            }
        }

        if (value != null)
        {
            RawValue = isTextRow
                ? sharedStrings[value]
                : value;
        }
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
            }

            _isDisposed = true;
        }
    }

    ~Cell()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }

    /// <InheritDoc />
    public object? RawValue { get; }

    /// <InheritDoc />
    public Type RawExcelType => throw new NotImplementedException();

    /// <InheritDoc />
    public ReadOnlyMemory<char> ColumnLetters { get; }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; }
}
