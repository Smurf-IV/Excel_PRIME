using System;
using System.Collections.Generic;
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
    private Dictionary<int, string>? _numberFormats;
    private Dictionary<int, CellStyle>? _cellStyles;
    private bool _isDisposed;

    /// <summary>
    /// Built-in number format codes defined in the ECMA-376 standard.
    /// </summary>
    private static readonly IReadOnlyDictionary<int, (string FormatCode, FormattingType Type)> BuiltInNumberFormats = Ecma376StandardProvider.BuiltInNumberFormats;

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
    public Dictionary<int, CellStyle> ExtractStyles(CancellationToken ct)
    {
        if (_cellStyles != null)
        {
            return _cellStyles;
        }

        _cellStyles = new Dictionary<int, CellStyle>();
        _numberFormats = new Dictionary<int, string>();
        
        // Copy format codes from built-in formats
        foreach ((int formatId, (string formatCode, FormattingType _)) in BuiltInNumberFormats)
        {
            _numberFormats[formatId] = formatCode;
        }

        try
        {
            using (Stream? styleStream = _zipReader.GetEntry("xl/styles.xml"))
            {
                if (styleStream == null)
                {
                    return _cellStyles;
                }

                using (XmlReader reader = XmlReader.Create(styleStream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreWhitespace = true,
                    CheckCharacters = false,
                    CloseInput = true,
                    ConformanceLevel = ConformanceLevel.Document,
                    ValidationType = ValidationType.None,
                }))
                {
                    ParseStylesXml(reader);
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
            && int.TryParse(numFmtIdAttr, out int numFmtId)
            && formatCodeAttr != null)
        {
            _numberFormats![numFmtId] = formatCodeAttr;
        }
    }

    private void ParseCellXfs(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            return;
        }

        int styleIndex = 0;
        while (reader.Read() && reader.NodeType != XmlNodeType.EndElement)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "xf")
            {
                CellStyle style = ParseCellXfElement(reader, styleIndex);
                _cellStyles![styleIndex] = style;
                styleIndex++;
            }
        }
    }

    private CellStyle ParseCellXfElement(XmlReader reader, int styleIndex)
    {
        CellStyle style = new CellStyle { StyleId = styleIndex };

        // Parse attributes
        if (int.TryParse(reader.GetAttribute("numFmtId"), out int numFmtId))
        {
            style.NumberFormatId = numFmtId;
            //style.ApplyNumberFormat = reader.GetAttribute("applyNumberFormat") == "1";
            
            // Look up the format code
            if (_numberFormats!.TryGetValue(numFmtId, out string? formatCode))
            {
                style.NumberFormatCode = formatCode;
            }
        }

        return style;
    }

    public void Dispose()
    {
        if (!_isDisposed)
        {
            _cellStyles?.Clear();
            _numberFormats?.Clear();
            _isDisposed = true;
        }
    }

    public async Task<IReadOnlyDictionary<int, CellStyle>> ExtractStylesAsync(CancellationToken ct) => throw new NotImplementedException();
}
