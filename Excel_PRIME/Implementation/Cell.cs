using System;
using System.Threading.Tasks;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class Cell : ICell
{
    private bool _isDisposed;
    /*
    | Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
    |-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
    | AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  42000.0000 |  40000.0000 |   5000.0000 |  334.78 MB |             |
    | AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.41x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB | 10.10x more |
    | AccessEveryCellFastExcel        | Data/100mb.xlsx      | 4.47x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB | 10.10x more |
    | AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 2.02x slower | 688000.0000 |  35000.0000 |   5000.0000 |  5464.9 MB | 16.32x more |
    | AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 2.06x slower | 680000.0000 |  35000.0000 |   5000.0000 | 5388.16 MB | 16.09x more |
    |                                 |                      |              |             |             |             |            |             |
    | AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 391000.0000 | 375000.0000 | 374000.0000 | 2696.79 MB |             |
    | AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.03x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
    | AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 1.01x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
    | AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.69x slower | 979000.0000 |   5000.0000 |   2000.0000 | 7814.39 MB |  2.90x more |
    | AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.81x slower | 971000.0000 |   4000.0000 |   2000.0000 | 7744.17 MB |  2.87x more |
    |                                 |                      |              |             |             |             |            |             |
    | AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  13000.0000 |           - |           - |  106.77 MB |             |
    | AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.04x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
    | AccessEveryCellFastExcel        | Data/(...).xlsx [39] | 1.05x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
    | AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.75x slower | 449000.0000 |   1000.0000 |           - | 3584.97 MB | 33.58x more |
    | AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.86x slower | 445000.0000 |   1000.0000 |           - | 3550.62 MB | 33.25x more |
    |                                 |                      |              |             |             |             |            |             |
    | AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  13000.0000 |           - |           - |  104.75 MB |             |
    | AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.08x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
    | AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 1.08x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
    | AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.87x slower | 442000.0000 |   1000.0000 |           - | 3526.56 MB | 33.67x more |
    | AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.94x slower | 438000.0000 |   1000.0000 |           - | 3492.16 MB | 33.34x more |
     */
    /*    public static async Task<Cell> ConstructCellAsync(XmlReader reader, ISharedString sharedStrings)
        {
            string? address = reader.GetAttribute("r");
            string? type = reader.GetAttribute("t");
            //var style = reader.GetAttribute("s");
            string? value = null;
            if (reader.ReadToDescendant("v") && !reader.IsEmptyElement)
            {
                value = await reader.ReadElementContentAsStringAsync().ConfigureAwait(false);
            }

            if (type == "s"
                && value != null)
            {   // Shared string
                value = sharedStrings[value];
            }

            // If this goes boom, then something is seriously wrong,
            // TODO: The exception needs to state something useful!
            (int _, int col, ReadOnlyMemory<char> colName) = address.GetRowColNumbers();
            return new Cell
            {
                ColumnLetters = colName,
                //RowNumber = row;
                ExcelColumnOffset = col,
                RawExcelType = value?.GetType(),
                RawValue = value
            };
        }
    */
    /*
| Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
|-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
| AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  42000.0000 |  40000.0000 |   5000.0000 |  334.79 MB |             |
| AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.38x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB | 10.10x more |
| AccessEveryCellFastExcel        | Data/100mb.xlsx      | 4.34x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB | 10.10x more |
| AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 1.77x slower | 457000.0000 |  34000.0000 |   5000.0000 | 3622.62 MB | 10.82x more |
| AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.78x slower | 448000.0000 |  34000.0000 |   5000.0000 | 3545.91 MB | 10.59x more |
|                                 |                      |              |             |             |             |            |             |
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 387000.0000 | 371000.0000 | 370000.0000 | 2696.72 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.04x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
| AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 1.05x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.60x slower | 572000.0000 |   2000.0000 |   1000.0000 | 4566.53 MB |  1.69x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.66x slower | 564000.0000 |   2000.0000 |   1000.0000 |  4496.4 MB |  1.67x more |
|                                 |                      |              |             |             |             |            |             |
| AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  13000.0000 |           - |           - |  106.77 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.04x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
| AccessEveryCellFastExcel        | Data/(...).xlsx [39] | 1.06x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 1.51x slower | 268000.0000 |           - |           - | 2141.68 MB | 20.06x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 1.66x slower | 264000.0000 |   1000.0000 |           - | 2107.34 MB | 19.74x more |
|                                 |                      |              |             |             |             |            |             |
| AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  13000.0000 |           - |           - |  104.75 MB |             |
| AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.11x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
| AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 1.09x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
| AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 1.63x slower | 261000.0000 |   1000.0000 |           - | 2082.97 MB | 19.89x more |
| AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 1.66x slower | 257000.0000 |   1000.0000 |           - |  2048.6 MB | 19.56x more |
    */
    public static async Task<Cell> ConstructCellAsync(XmlReader reader, ISharedString sharedStrings)
    {
        string address = string.Empty;
        string type = string.Empty;
        while (reader.MoveToNextAttribute())
        {
            switch (reader.LocalName)
            {
                case "r":
                    address = reader.Value;
                    break;
                case "t":
                    type = reader.Value;
                    break;
            }
        }
        string? value = null;
        if (await reader.ReadAsync().ConfigureAwait(false)
            && reader is { IsEmptyElement: false, LocalName: "v" })
        {
            if (reader.IsStartElement()
                || await reader.MoveToContentAsync().ConfigureAwait(false) == XmlNodeType.Element
               )
            {
                value = reader.ReadString();
                if (type == "s")
                {   // Shared string
                    value = sharedStrings[value];
                }
            }
        }

        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        (int _, int col, ReadOnlyMemory<char> colName) = address.GetRowColNumbers();
        return new Cell
        {
            ColumnLetters = colName,
            //RowNumber = row;
            ExcelColumnOffset = col,
            RawExcelType = value?.GetType(),
            RawValue = value
        };
    }


    /*
    | Method                          | FileName             | Ratio        | Gen0        | Gen1        | Gen2        | Allocated  | Alloc Ratio |
    |-------------------------------- |--------------------- |-------------:|------------:|------------:|------------:|-----------:|------------:|
    | AccessEveryCellSylvan           | Data/100mb.xlsx      |     baseline |  42000.0000 |  40000.0000 |   5000.0000 |   334.8 MB |             |
    | AccessEveryCellXlsxHelper       | Data/100mb.xlsx      | 4.24x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.59 MB | 10.10x more |
    | AccessEveryCellFastExcel        | Data/100mb.xlsx      | 4.37x slower | 424000.0000 |   5000.0000 |   2000.0000 | 3380.58 MB | 10.10x more |
    | AccessEveryCellAsyncExcel_Prime | Data/100mb.xlsx      | 2.03x slower | 615000.0000 |  35000.0000 |   5000.0000 | 4875.38 MB | 14.56x more |
    | AccessEveryCellExcel_Prime      | Data/100mb.xlsx      | 1.99x slower | 606000.0000 |  36000.0000 |   5000.0000 | 4798.68 MB | 14.33x more |
    |                                 |                      |              |             |             |             |            |             |
    | AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline | 390000.0000 | 374000.0000 | 373000.0000 | 2696.74 MB |             |
    | AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.03x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
    | AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 1.01x slower | 218000.0000 |   1000.0000 |           - | 1739.24 MB |  1.55x less |
    | AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] |            ? |          NA |          NA |          NA |         NA |           ? |
    | AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] |            ? |          NA |          NA |          NA |         NA |           ? |
    |                                 |                      |              |             |             |             |            |             |
    | AccessEveryCellSylvan           | Data/(...).xlsx [39] |     baseline |  13000.0000 |           - |           - |  106.77 MB |             |
    | AccessEveryCellXlsxHelper       | Data/(...).xlsx [39] | 1.04x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
    | AccessEveryCellFastExcel        | Data/(...).xlsx [39] | 1.05x slower | 100000.0000 |           - |           - |  799.73 MB |  7.49x more |
    | AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [39] | 2.00x slower | 544000.0000 |   1000.0000 |           - |  4340.8 MB | 40.65x more |
    | AccessEveryCellExcel_Prime      | Data/(...).xlsx [39] | 2.10x slower | 540000.0000 |   1000.0000 |           - | 4306.44 MB | 40.33x more |
    |                                 |                      |              |             |             |             |            |             |
    | AccessEveryCellSylvan           | Data/(...).xlsx [35] |     baseline |  13000.0000 |           - |           - |  104.75 MB |             |
    | AccessEveryCellXlsxHelper       | Data/(...).xlsx [35] | 1.09x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
    | AccessEveryCellFastExcel        | Data/(...).xlsx [35] | 1.12x slower |  93000.0000 |           - |           - |  742.13 MB |  7.08x more |
    | AccessEveryCellAsyncExcel_Prime | Data/(...).xlsx [35] | 2.03x slower | 546000.0000 |   1000.0000 |           - | 4362.47 MB | 41.65x more |
    | AccessEveryCellExcel_Prime      | Data/(...).xlsx [35] | 2.10x slower | 543000.0000 |   1000.0000 |           - | 4328.07 MB | 41.32x more |
     */
    /*public static async Task<Cell> ConstructCellAsync(XmlReader reader, ISharedString sharedStrings)
    {
        char[] addressBuf = new char[16];
        char[] type = new char[16];
        //string type = string.Empty;// = reader.GetAttribute("t");
        int len = 0;
        int typeLen = 0;
        char[] buffer = new char[16];
        while (reader.MoveToNextAttribute())
        {
            switch (reader.LocalName[0])
            {
                case 'r':
                    len = await reader.ReadValueChunkAsync(addressBuf, 0, addressBuf.Length).ConfigureAwait(false);
                    //address = addressBuf;
                    break;
                case 't':
                    typeLen = await reader.ReadValueChunkAsync(type, 0, type.Length).ConfigureAwait(false);
                    //type = reader.Value;
                    break;
                case 's':
                    await reader.ReadValueChunkAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                    //if (!TryParse(buffer.AsSpan().ToParsable(0, len), out xfIdx))
                    //{
                    //    throw new FormatException();
                    //}
                    break;
            }
        }

        object? value = null;
        if (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (!reader.IsEmptyElement)
            {
                if (reader.IsStartElement()
                    || await reader.MoveToContentAsync().ConfigureAwait(false) == XmlNodeType.Element
                   )
                {
                    switch (type[0])
                    {
                        case 's':
                            if (typeLen == 1)
                            {
                                value = sharedStrings[reader.ReadString()];
                            }
                            else
                            {
                                string s = await reader.ReadElementContentAsStringAsync().ConfigureAwait(false);
                                if (reader.XmlSpace != XmlSpace.Preserve)
                                {
                                    s = s.Trim();
                                }
                                value = s;
                            }
                            break;

                        //case 'i': // Actually "Inline string" = `is` probably
                        //    value = reader.ReadString();
                        //    break;
                        //case 'b':
                        //    value = reader.ReadElementContentAsBoolean();
                        //    break;
                        //case 'd':
                        //    value = reader.ReadElementContentAsDateTime();
                        //    break;
                        //case 'n':
                        //// TODO: this is where styles could help
                        //value = await reader.ReadElementContentAsObjectAsync().ConfigureAwait(false);    // <- will default to string
                        //break;
                        default:
                            value = await reader.ReadElementContentAsObjectAsync().ConfigureAwait(false);    // <- will default to string
                            break;
                        case 'e':
                            // It's an error
                            value = reader.ReadString();
                            break;
                        //default:
                        //    // TODO: What is it now ?
                        //    value = await reader.ReadElementContentAsObjectAsync().ConfigureAwait(false);    // <- will default to string
                        //    if (typeLen > 0)
                        //        throw new IndexOutOfRangeException(new string(type));
                        //    break;
                    }
                }
            }
        }
        string address = new string(addressBuf, 0, len);
        // If this goes boom, then something is seriously wrong,
        // TODO: The exception needs to state something useful!
        (int _, int col, ReadOnlyMemory<char> colName) = address.GetRowColNumbers();
        return new Cell
        {
            ColumnLetters = colName,
            //RowNumber = row;
            ExcelColumnOffset = col,
            RawExcelType = value?.GetType(),
            RawValue = value
        };

    }
*/

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
    public object? RawValue { get; init; }

    /// <InheritDoc />
    public Type? RawExcelType { get; init; }

    /// <InheritDoc />
    public ReadOnlyMemory<char> ColumnLetters { get; init; }

    /// <InheritDoc />
    public int ExcelColumnOffset { get; init; }
}
