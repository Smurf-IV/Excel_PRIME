using ExcelPRIME.XlsbImp;


namespace ExcelPRIME;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores
/// <summary>
/// Main entry point into the `Excel_PRIMEXlsb` API's
/// </summary>
public class Excel_PRIMEXlsb : Excel_PRIME
{
    /// <InheritDoc />
    public Excel_PRIMEXlsb(IOpenXmlReaderHelpers? xlsbReader = null, IZipReaderAsync? zipReader = null)
    : base(xlsbReader ?? new XlsbReaderHelpers(), zipReader)
    {
    }
}

#pragma warning disable CA1707 // Underscores
/// <summary>
/// Main entry point into the `Excel_PRIMEXlsb` Async API's
/// </summary>
public sealed class Excel_PRIMEXlsbAsync : Excel_PRIMEAsync, IExcel_PRIMEXlsbAsync
{
    /// <InheritDoc />
    public Excel_PRIMEXlsbAsync(IOpenXmlReaderHelpersAsync? xlsbReader = null, IZipReaderAsync? zipReader = null)
        : base(xlsbReader ?? new XlsbReaderHelpersAsync(), zipReader)
    {
    }
}
