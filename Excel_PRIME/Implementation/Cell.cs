using System;
using System.Linq;
using System.Xml.Linq;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class Cell : ICell
{
    private bool _isDisposed;

    public Cell(XElement cellElement, ISharedString sharedStrings)
    {
        bool isTextRow = cellElement.Attributes("t").Any(a => a.Value == "s");
        string? columnName = cellElement.Attributes("r").Select(a => a.Value).FirstOrDefault();

        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        ColumnLetters = ExcelColumns.RemoveNumbers().Replace(columnName, string.Empty);

        //RowNumber = Convert.ToInt32(Regex.Replace(columnName, @"[^\d]", ""));

        //ColumnName = worksheet.FastExcel.DefinedNames.FindColumnName(worksheet.Name, columnLetter) ?? columnLetter;

        //CellNames = worksheet.FastExcel.DefinedNames.FindCellNames(worksheet.Name, columnLetter, RowNumber);

        ExcelColumnOffset = ColumnLetters.GetExcelColumnNumber(false);

        if (isTextRow)
        {
            RawValue = sharedStrings[cellElement.Value];
        }
        else
        {
            // cellElement.Value will give a concatenated Value + reference/calculation
            XElement? node = cellElement.Elements()
                .SingleOrDefault(x => x.Name.LocalName == "v");

            if (node != null)
            {
                RawValue = node.Value;
            }
            else
            {
                RawValue = cellElement.Value;
            }

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
    public string ColumnLetters { get; }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; }
}
