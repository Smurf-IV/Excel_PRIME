using ExcelFAST.Implementation;
using ExcelFAST.XlsbImp;

using ExcelPRIME;
using ExcelPRIME.Implementation;
using ExcelPRIME.XlsbImp;


namespace ExcelFAST;

/// <summary>
/// Main entry point into the `Excel_PRIMEXlsb` API's
/// </summary>
public sealed class ExcelFastXlsb : ExcelFastBase, IExcelFast
{
    /// <InheritDoc />
    public ExcelFastXlsb(IOpenXmlReaderHelpers? xlsbReader = null, IZipReader? zipReader = null)
    : base(xlsbReader ?? new XlsbReaderHelpers(), 
            zipReader ?? new ZipReader(),
            new RowFastXlsb()
            )
    {
    }
}
