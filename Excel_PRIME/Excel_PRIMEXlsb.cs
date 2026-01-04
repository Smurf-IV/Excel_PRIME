using ExcelPRIME.XlsbImp;


namespace ExcelPRIME;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
/// <summary>
/// Main entry point into the `Excel_PRIMEXlsb` API's
/// </summary>
public sealed class Excel_PRIMEXlsb : Excel_PRIME
{
    /// <InheritDoc />
    public Excel_PRIMEXlsb(IOpenXmlReaderHelpersAsync? xlsbReader = null, IZipReaderAsync? zipReader = null)
    : base(xlsbReader ?? new XlsbReaderHelpersAsync(), zipReader)
    {
    }
}
