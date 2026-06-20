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
    private readonly Dictionary<short, CellStyle> _numberFormats = Ecma376StandardProvider.GetDefaultStyles();
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
            XmlDocument doc = new();
            doc.Load(styleStream);
            if (doc.DocumentElement == null)
            {
                throw new InvalidDataException();
            }
            XmlNamespaceManager nsm = new(doc.NameTable);
            string ns = doc.DocumentElement.NamespaceURI;
            nsm.AddNamespace("x", ns);
            XmlNodeList? nodes = doc.SelectNodes("/x:styleSheet/x:numFmts/x:numFmt", nsm);

            if (nodes != null)
            {
                foreach (XmlElement fmt in nodes)
                {
                    ParseNumFmt(fmt);
                }
            }
            XmlElement? xfsElem = (XmlElement?)doc.SelectSingleNode("/x:styleSheet/x:cellXfs", nsm);
            if (xfsElem != null)
            {
                IEnumerable<XmlElement> cellNodes = xfsElem.ChildNodes.OfType<XmlElement>();
                short styleIndex = 0;

                foreach ( XmlElement cellNode in cellNodes)
                {
                    _cellStyles[styleIndex++] = ParseCellXfElement(cellNode);
                }
            }
        }
        catch (Exception)
        {
            // If style extraction fails, return dictionary with default styles
            // This allows the workbook to continue working without style information
        }

        return _cellStyles;
    }


    public async Task<IReadOnlyDictionary<short, CellStyle>> ExtractStylesAsync(CancellationToken ct)
    {
        try
        {
            using Stream? styleStream = await ((IZipReaderAsync)_zipReader).GetEntryAsync("xl/styles.xml", ct).ConfigureAwait(false);
            if (styleStream == null)
            {
                return _cellStyles;
            }

            using XmlReader reader = XmlReader.Create(styleStream, new XmlReaderSettings
            {
                Async = true,
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                CheckCharacters = false
            });

            string numFmtsAtom = reader.NameTable.Add("numFmts");
            string numFmtAtom = reader.NameTable.Add("numFmt");
            string cellXfsAtom = reader.NameTable.Add("cellXfs");
            string xfAtom = reader.NameTable.Add("xf");

            while (await reader.ReadAsync().ConfigureAwait(false) && !ct.IsCancellationRequested)
            {
                if (reader.NodeType != XmlNodeType.Element)
                {
                    continue;
                }

                if (ReferenceEquals(reader.LocalName, numFmtsAtom))
                {
                    if (reader.IsEmptyElement)
                    {
                        continue;
                    }

                    while (await reader.ReadAsync().ConfigureAwait(false) && !ct.IsCancellationRequested)
                    {
                        if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, numFmtAtom))
                        {
                            ParseNumFmt(reader);
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, numFmtsAtom))
                        {
                            break;
                        }
                    }
                }
                else if (ReferenceEquals(reader.LocalName, cellXfsAtom))
                {
                    if (reader.IsEmptyElement)
                    {
                        continue;
                    }

                    short styleIndex = 0;
                    while (await reader.ReadAsync().ConfigureAwait(false) && !ct.IsCancellationRequested)
                    {
                        if (reader.NodeType == XmlNodeType.Element && ReferenceEquals(reader.LocalName, xfAtom))
                        {
                            _cellStyles[styleIndex++] = ParseCellXfElement(reader);
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement && ReferenceEquals(reader.LocalName, cellXfsAtom))
                        {
                            break;
                        }
                    }
                }
            }
        }
        catch
        {
            // Do nothing
        }

        return _cellStyles;
    }

    private void ParseNumFmt(XmlReader reader)
    {
        string numFmtIdAttr = reader.GetAttribute("numFmtId") ?? string.Empty;
        string formatCodeAttr = reader.GetAttribute("formatCode") ?? string.Empty;

        if (short.TryParse(numFmtIdAttr, out short numFmtId)
            && !string.IsNullOrWhiteSpace(formatCodeAttr)
            )
        {
            _numberFormats[numFmtId] = new CellStyle
            {
                ExcelFormatId = numFmtId,
                Formatting = formatCodeAttr,
            };
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
            _numberFormats.TryGetValue(numFmtId, out style);
            style ??= Ecma376StandardProvider.GetCellStyle(0);
        }

        return style!;
    }

    private void ParseNumFmt(XmlElement fmt)
    {
        string numFmtIdAttr = fmt.GetAttribute("numFmtId");
        string formatCodeAttr = fmt.GetAttribute("formatCode");

        if (!string.IsNullOrWhiteSpace(numFmtIdAttr) 
            && short.TryParse(numFmtIdAttr, out short numFmtId)
            && !string.IsNullOrWhiteSpace(formatCodeAttr))
        {
            _numberFormats[numFmtId] = new CellStyle
            {
                ExcelFormatId = numFmtId,
                Formatting = formatCodeAttr,
                //FormattingType = // TODO: check if it contains magic for dates (i,e, locale stuff as well!)
            };
        }
    }


    private CellStyle ParseCellXfElement(XmlElement reader)
    {
        CellStyle? style = null;
        if (short.TryParse(reader.GetAttribute("numFmtId"), out short numFmtId))
        {
            Ecma376StandardProvider.TryGetCellStyle(numFmtId, out style);
        }
        if (style == null)
        {
            _numberFormats.TryGetValue(numFmtId, out style);
            style ??= Ecma376StandardProvider.GetCellStyle(0);
        }

        return style!;
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
}
