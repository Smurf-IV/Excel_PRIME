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
    private readonly Dictionary<short, CellStyle> _numberFormats = Ecma376StandardProvider.GetDefaultStyles();
    private readonly Dictionary<short, CellStyle> _cellStyles = [];
    private bool _isDisposed;

    /// <summary>
    /// XLSB Record Type IDs for styles-related records.
    /// These are based on the ECMA-376 Part 2 specification for XLSB format.
    /// </summary>
    private static class XlsbRecordTypes
    {
        public const int NumFmtStart = 615;
        public const int NumFmt = 44;

        public const int CellXFStart = 617;
        public const int CellXF = 47;
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
    public IReadOnlyDictionary<short, CellStyle> ExtractStyles(CancellationToken ct)
    {
        // Add XLSB default/implicit styles per ECMA-376 specification
        AddDefaultStyles();

        try
        {
            using Stream? styleStream = _zipReader.GetEntry("xl/styles.bin");
            if (styleStream == null)
            {
                return _cellStyles;
            }

            using BufferedStream bufferedStream = new(styleStream, 64 * 1024);
            XlsbStreamReader reader = new(bufferedStream);
            ParseStylesXlsb(reader);
            // XlsbStreamReader doesn't own the stream, so no dispose needed
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
    /// Reference: ECMA-376-1:2016 Section 18.8.10 (Cell Formats - cellXfs)
    /// </summary>
    private void AddDefaultStyles()
    {
        _cellStyles[0] = Ecma376StandardProvider.GetCellStyle(0)!;
        _cellStyles[1] = Ecma376StandardProvider.GetCellStyle(3)!;// Style 1 - Comma format
        _cellStyles[2] = Ecma376StandardProvider.GetCellStyle(4)!;// Style 2 - Comma (2 decimal places)
        _cellStyles[3] = Ecma376StandardProvider.GetCellStyle(5)!;// Style 3 - Currency
        _cellStyles[4] = Ecma376StandardProvider.GetCellStyle(6)!;// Style 4 - Currency (2 decimal places)
        _cellStyles[5] = Ecma376StandardProvider.GetCellStyle(90)!;// Style 5 - Percent
    }

    private void ParseStylesXlsb(XlsbStreamReader reader)
    {
        PooledRecordBuffer record = reader.ReadNextRecord();
        bool exitNow = false;
        while (!exitNow && record.Succeeded)
        {
            try
            {
                short count;
                switch (record.RecordType)
                {
                    case (RecordTypeIdentifier)XlsbRecordTypes.NumFmtStart:
                        count = record.GetInt16(0);
                        for (short offset = 0; offset < count; offset++)
                        {
                            record.Dispose();
                            record = reader.ReadNextRecord();
                            if (record.RecordType != (RecordTypeIdentifier)XlsbRecordTypes.NumFmt)
                            {
                                throw new InvalidDataException();
                            }
                            ParseNumFmt(record, count);
                        }
                        break;

                    case (RecordTypeIdentifier)XlsbRecordTypes.CellXFStart:
                        count = record.GetInt16(0);
                        for (short offset = 0; offset < count; offset++)
                        {
                            record.Dispose();
                            record = reader.ReadNextRecord();
                            if (record.RecordType != (RecordTypeIdentifier)XlsbRecordTypes.CellXF)
                            {
                                throw new InvalidDataException();
                            }
                            ParseCellXf(record, offset);
                        }
                        exitNow = true;
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
    private void ParseNumFmt(PooledRecordBuffer record, short count)
    {
        try
        {
            short numFmtId = record.GetInt16(0);
            string? formatCode = ParseStringFromRecord(record, 2);

            _numberFormats[numFmtId] = new CellStyle
            {
                ExcelFormatId = numFmtId,
                Formatting = formatCode,
                //FormattingType = // TODO: check if it contains magic for dates (i,e, locale stuff as well!)
            };
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
    private void ParseCellXf(PooledRecordBuffer record, short styleIndex)
    {
        try
        {
            CellStyle? style = null;
            // Extract IDs from the record
            // Offsets are based on XLSB CELLXF record format
            if (TryGetInt16(record, 2, out short numFmtId))
            {
                Ecma376StandardProvider.TryGetCellStyle(numFmtId, out style);
            }

            if (style == null)
            {
                _numberFormats.TryGetValue(numFmtId, out style);
                style ??= Ecma376StandardProvider.GetCellStyle(0);
            }

            _cellStyles[styleIndex] = style!;
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
            // Use span-based GetString to avoid intermediate byte array allocation
            return Encoding.Unicode.GetString(record.AsSpan(offset + 4, length * 2)).TrimEnd('\0');
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
            // _cellStyles.Clear(); returned to the caller, so we should not clear it here
            _numberFormats.Clear();
            _isDisposed = true;
        }
    }

    public async Task<IReadOnlyDictionary<short, CellStyle>> ExtractStylesAsync(CancellationToken ct)
    {
        // Add XLSB default/implicit styles per ECMA-376 specification
        AddDefaultStyles();

        try
        {
            using Stream? styleStream = await ((IZipReaderAsync)_zipReader).GetEntryAsync("xl/styles.bin", ct).ConfigureAwait(false);
            if (styleStream == null)
            {
                return _cellStyles;
            }

            using BufferedStream bufferedStream = new(styleStream, 64 * 1024);
            XlsbStreamReader reader = new(bufferedStream);
            await ParseStylesXlsbAsync(reader, ct).ConfigureAwait(false);
            // XlsbStreamReader doesn't own the stream, so no dispose needed
        }
        catch (Exception)
        {
            // If style extraction fails, return dictionary with default styles
            // This allows the workbook to continue working without style information
        }

        return _cellStyles;
    }

    private async Task ParseStylesXlsbAsync(XlsbStreamReader reader, CancellationToken ct)
    {
        PooledRecordBuffer record = await reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
        bool exitNow = false;
        while (!exitNow && record.Succeeded && !ct.IsCancellationRequested)
        {
            try
            {
                short count;
                switch (record.RecordType)
                {
                    case (RecordTypeIdentifier)XlsbRecordTypes.NumFmtStart:
                        count = record.GetInt16(0);
                        for (short offset = 0; offset < count; offset++)
                        {
                            record.Dispose();
                            record = await reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
                            if (record.RecordType != (RecordTypeIdentifier)XlsbRecordTypes.NumFmt)
                            {
                                throw new InvalidDataException();
                            }
                            ParseNumFmt(record, count);
                        }
                        break;

                    case (RecordTypeIdentifier)XlsbRecordTypes.CellXFStart:
                        count = record.GetInt16(0);
                        for (short offset = 0; offset < count; offset++)
                        {
                            record.Dispose();
                            record = await reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
                            if (record.RecordType != (RecordTypeIdentifier)XlsbRecordTypes.CellXF)
                            {
                                throw new InvalidDataException();
                            }
                            ParseCellXf(record, offset);
                        }
                        exitNow = true;
                        break;
                }
            }
            finally
            {
                record.Dispose();
            }

            if (!exitNow)
            {
                record = await reader.ReadNextRecordAsync(ct).ConfigureAwait(false);
            }
        }
    }
}
