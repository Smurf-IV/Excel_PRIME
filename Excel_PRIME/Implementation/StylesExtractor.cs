using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelPRIME.Implementation;

/// <summary>
/// Extracts cell styling information from an XLSX workbook's styles.xml file.
/// </summary>
internal sealed class StylesExtractor : IDisposable
{
    private readonly IZipReader _zipReader;
    private readonly Dictionary<short, string> _numberFormats = [];
    private readonly Dictionary<short, CellStyle> _cellStyles = [];
    private bool _isDisposed;

    public StylesExtractor(IZipReader zipReader)
    {
        _zipReader = zipReader ?? throw new ArgumentNullException(nameof(zipReader));
    }

    /// <summary>
    /// Extracts all cell styles from the workbook's styles.xml file.
    /// </summary>
    /// <param name="ct"></param>
    /// <remarks>
    /// Per ECMA-376 specification, default/implicit styles are pre-populated with standard IDs.
    /// </remarks>
    /// <returns>
    /// A dictionary mapping style IDs to CellStyle objects. Always includes built-in default styles.
    /// Returns empty/default styles if styles.xml is not found.
    /// </returns>
    public IReadOnlyDictionary<short, CellStyle> ExtractStyles(CancellationToken ct)
    {
        try
        {
            using Stream? styleStream = _zipReader.GetEntry("xl/styles.xml");
            if (styleStream == null)
            {
                return _cellStyles;
            }

            using XmlReader reader = XmlReader.Create(styleStream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                CheckCharacters = false,
                CloseInput = true,
                ConformanceLevel = ConformanceLevel.Document,
                ValidationType = ValidationType.None
            });
            ParseStylesXml(reader);
        }
        catch (Exception)
        {
            // If style extraction fails, return dictionary with default styles
            // This allows the workbook to continue working without style information
        }

        return _cellStyles;
    }

    private void ParseStylesXml(XmlReader reader)
    {
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case "numFmt":
                        ParseNumFmt(reader);
                        break;
                    case "cellXfs":
                        ParseCellXfs(reader);
                        break;
                }
            }
        }
    }

    private void ParseNumFmt(XmlReader reader)
    {
        string? numFmtIdAttr = reader.GetAttribute("numFmtId");
        string? formatCodeAttr = reader.GetAttribute("formatCode");

        if (numFmtIdAttr != null
            && short.TryParse(numFmtIdAttr, out short numFmtId)
            && formatCodeAttr != null)
        {
            _numberFormats[numFmtId] = formatCodeAttr;
        }
    }

    private void ParseCellXfs(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            return;
        }

        short styleIndex = 0;
        while (reader.Read() && reader.NodeType != XmlNodeType.EndElement)
        {
            if (reader is { NodeType: XmlNodeType.Element, LocalName: "xf" })
            {
                CellStyle style = ParseCellXfElement(reader);
                _cellStyles[styleIndex] = style;
                styleIndex++;
            }
        }
    }

    private CellStyle ParseCellXfElement(XmlReader reader)
    {
        CellStyle? style = null;
        if (short.TryParse(reader.GetAttribute("numFmtId"), out short numFmtId))
        {
            Ecma376StandardProvider.TryGetCellStyle(numFmtId, out style);
        }
        if (style == null)
        {
            // Parse attributes
            //style.ApplyNumberFormat = reader.GetAttribute("applyNumberFormat") == "1";
            _numberFormats.TryGetValue(numFmtId, out string? formatCode);
            style = new CellStyle
            {
                ExcelFormatId = numFmtId,
                Formatting = formatCode
            };
        }

        return style;
    }

    public void Dispose()
    {
        if (!_isDisposed)
        {
            //_cellStyles?.Clear(); returned to caller
            _numberFormats.Clear();
            _isDisposed = true;
        }
    }

    public Task<IReadOnlyDictionary<short, CellStyle>> ExtractStylesAsync(CancellationToken ct)
        => Task.FromResult(ExtractStyles(ct));
}
