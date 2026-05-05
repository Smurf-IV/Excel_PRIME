using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.Implementation;

namespace ExcelPRIME.XlsbImp;

/// <summary>
/// Extracts cell styling information from an XLSB workbook's styles.bin file.
/// </summary>
internal sealed class XlsbStylesExtractor : IDisposable
{
    private readonly IZipReader _zipReader;
    private Dictionary<int, string>? _numberFormats;
    private Dictionary<int, CellStyle>? _cellStyles;
    private bool _isDisposed;

    /// <summary>
    /// Built-in number format codes defined in the ECMA-376 standard.
    /// </summary>
    private static readonly IReadOnlyDictionary<int, (string FormatCode, FormattingType Type)> BuiltInNumberFormats = Ecma376StandardProvider.BuiltInNumberFormats;

    /// <summary>
    /// XLSB Record Type IDs for styles-related records.
    /// These are based on the ECMA-376 Part 2 specification for XLSB format.
    /// </summary>
    private static class XlsbRecordTypes
    {
        // Number Formats
        public const int NUMFMT = 0x0450; // 1104

        // Fonts
        public const int FONT = 0x0470; // 1136

        // Fills
        public const int FILL = 0x002F; // 47

        // Borders
        public const int BORDER = 0x0471; // 1137

        // Alignments
        public const int XFID = 0x0408; // 1032

        // Cell Format Style (cellXfs equivalent)
        public const int CELLXF = 0x044F; // 1103

        // Style Format (cellStyleXfs equivalent)
        public const int STYLEXF = 0x04AD; // 1197

        // Number Format ID (fmtId)
        public const int FMTID = 0x0449; // 1097
    }

    public XlsbStylesExtractor(IZipReader zipReader)
    {
        _zipReader = zipReader ?? throw new ArgumentNullException(nameof(zipReader));
    }

    /// <summary>
    /// Extracts all cell styles from the workbook's styles.bin file.
    /// </summary>
    /// <remarks>
    /// Per ECMA-376 specification, default/implicit styles are pre-populated with standard IDs.
    /// </remarks>
    /// <returns>
    /// A dictionary mapping style IDs to CellStyle objects. Always includes built-in default styles.
    /// Returns empty/default styles if styles.bin is not found.
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

        // Add XLSB default/implicit styles per ECMA-376 specification
        AddDefaultStyles();

        try
        {
            using (Stream? styleStream = _zipReader.GetEntry("xl/styles.bin"))
            {
                if (styleStream == null)
                {
                    return _cellStyles;
                }

                using (BufferedStream bufferedStream = new(styleStream, 64 * 1024))
                {
                    XlsbStreamReader reader = new(bufferedStream);
                    ParseStylesXlsb(reader);
                    // XlsbStreamReader doesn't own the stream, so no dispose needed
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

    /// <summary>
    /// Adds the XLSB default/implicit styles per ECMA-376 Part 2.
    /// These are always present in an XLSB file, even if not explicitly defined in styles.bin.
    /// </summary>
    private void AddDefaultStyles()
    {
        IReadOnlyDictionary<int, CellStyle> defaultStyles = Ecma376StandardProvider.DefaultStyles.GetAll();
        foreach (KeyValuePair<int, CellStyle> kvp in defaultStyles)
        {
            _cellStyles![kvp.Key] = kvp.Value;
        }
    }

    private void ParseStylesXlsb(XlsbStreamReader reader)
    {
        PooledRecordBuffer record = reader.ReadNextRecord();
        int styleIndex = 0;

        while (record.Succeeded)
        {
            try
            {
                switch (record.RecordType)
                {
                    case (RecordTypeIdentifier)XlsbRecordTypes.NUMFMT:
                        ParseNumFmt(record);
                        break;

                    case (RecordTypeIdentifier)XlsbRecordTypes.CELLXF:
                        ParseCellXf(record, styleIndex);
                        styleIndex++;
                        break;
                }
            }
            finally
            {
                record.Dispose();
            }

            record = reader.ReadNextRecord();
        }
        record.Dispose();
    }

    /// <summary>
    /// Parses a NUMFMT record to extract custom number format codes.
    /// NUMFMT record structure:
    /// - numFmtId (4 bytes): Format ID
    /// - formatCode (variable): Format code string
    /// </summary>
    private void ParseNumFmt(PooledRecordBuffer record)
    {
        try
        {
            int numFmtId = record.GetInt32(0);
            // The format code string starts at offset 4
            string? formatCode = ParseStringFromRecord(record, 4);

            if (formatCode != null)
            {
                _numberFormats![numFmtId] = formatCode;
            }
        }
        catch (Exception)
        {
            // Gracefully handle parsing errors
        }
    }

    /// <summary>
    /// Parses a CELLXF (Cell Format) record.
    /// CELLXF record structure (simplified):
    /// - fontId (2 bytes): Font reference
    /// - numFmtId (2 bytes): Number format reference
    /// - fillId (2 bytes): Fill reference
    /// - borderId (2 bytes): Border reference
    /// - xfFormatId (2 bytes): Style XF reference
    /// - alignment info: Various alignment flags
    /// - protection info: Various protection flags
    /// </summary>
    private void ParseCellXf(PooledRecordBuffer record, int styleIndex)
    {
        try
        {
            CellStyle style = new CellStyle { StyleId = styleIndex };

            // Extract IDs from the record
            // Offsets are based on XLSB CELLXF record format
            if (record.RecordType == (RecordTypeIdentifier)XlsbRecordTypes.CELLXF)
            {
                // Note: Exact byte offsets depend on the XLSB specification
                // These are approximate based on common XLSB implementations
                // Offset 0-1: fontId
                // Offset 2-3: numFmtId
                if (TryGetInt16(record, 2, out short numFmtId))
                {
                    style.NumberFormatId = numFmtId;
                    //style.ApplyNumberFormat = true;

                    // Look up the format code
                    if (_numberFormats!.TryGetValue(numFmtId, out string? formatCode))
                    {
                        style.NumberFormatCode = formatCode;
                    }
                }
            }

            _cellStyles![styleIndex] = style;
        }
        catch (Exception)
        {
            // Gracefully handle parsing errors
        }
    }

    /// <summary>
    /// Tries to get a 16-bit integer from the record at the specified offset.
    /// </summary>
    private static bool TryGetInt16(PooledRecordBuffer record, int offset, out short value)
    {
        try
        {
            value = record.GetInt16(offset);
            return true;
        }
        catch
        {
            value = 0;
            return false;
        }
    }

    /// <summary>
    /// Tries to get a single byte from the record at the specified offset.
    /// </summary>
    private static bool TryGetByte(PooledRecordBuffer record, int offset, out byte value)
    {
        try
        {
            value = record.GetByte(offset);
            return true;
        }
        catch
        {
            value = 0;
            return false;
        }
    }

    /// <summary>
    /// Parses a UTF-16 LE string from the record at the specified offset.
    /// String format: 4-byte length (in characters) + character data
    /// </summary>
    private static string? ParseStringFromRecord(PooledRecordBuffer record, int offset)
    {
        try
        {
            // Read the string length (stored as 4-byte integer)
            int length = record.GetInt32(offset);

            if (length <= 0)
            {
                return string.Empty;
            }

            // The actual string data starts at offset + 4
            // In XLSB, strings are UTF-16 LE encoded
            byte[] buffer = new byte[length * 2];
            for (int i = 0; i < length * 2; i++)
            {
                buffer[i] = record.GetByte(offset + 4 + i);
            }

            return Encoding.Unicode.GetString(buffer).TrimEnd('\0');
        }
        catch
        {
            return null;
        }
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
