using ExcelFAST.Implementation;
using ExcelFAST.XlsxImp;

using ExcelPRIME;
using ExcelPRIME.Implementation;


namespace ExcelFAST;

/// <summary>
/// Main entry point into the `Excel_PRIME` XLSX handler API's
/// </summary>
public sealed class ExcelFastXlsx : ExcelFastBase, IExcelFast
{
    /// <InheritDoc />
    public ExcelFastXlsx(IOpenXmlReaderHelpers? xmlReader = null, IZipReader? zipReader = null)
    : base(xmlReader ?? new XmlReaderHelpers(),
            zipReader ?? new ZipReader(),
            new RowFastXlsx()
            )
    {
    }

    /// <inheritdoc />
    protected override IRowFast CreateRowFast() => new RowFastXlsx();
}