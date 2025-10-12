using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class Cell : ICell
{
    private bool _isDisposed;

    public Cell(XElement cellElement, IReadOnlyList<string> sharedStrings)
    {
        bool isTextRow = cellElement.Attributes("t").Any(a => a.Value == "s");
        string? columnName = cellElement.Attributes("r").Select(a => a.Value).FirstOrDefault();

        ColumnLetters = Regex.Replace(columnName, @"\d", "");

        //RowNumber = Convert.ToInt32(Regex.Replace(columnName, @"[^\d]", ""));

        //ColumnName = worksheet.FastExcel.DefinedNames.FindColumnName(worksheet.Name, columnLetter) ?? columnLetter;

        //CellNames = worksheet.FastExcel.DefinedNames.FindCellNames(worksheet.Name, columnLetter, RowNumber);

        ExcelColumnOffset = ColumnLetters.GetExcelColumnNumber();

        if (isTextRow)
        {
            RawValue = sharedStrings[int.Parse(cellElement.Value, CultureInfo.InvariantCulture)];
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
