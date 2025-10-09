using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

using Excel_PRIME.Implementation;
using Excel_PRIME.Shared;
using Excel_PRIME.XML;

// ReSharper disable InconsistentNaming
#pragma warning disable CA1707 // Underscores

namespace Excel_PRIME;

public class Excel_PRIME : IExcel_PRIME
{
    private bool _isDisposed;
    private readonly IXmlReaderHelpers _xmlReader;
    private readonly IZipReader _zipReader;
    private Stream? _fs;
    private readonly Dictionary<string, TempFile> _baseFiles = new();
    private readonly Dictionary<int, Sheet> _sheets = new();
    private ReadOnlyDictionary<string, int> _sheetNamesWithrId = new Dictionary<string, int>(0).AsReadOnly();

    public Excel_PRIME(IXmlReaderHelpers? xmlReader = null, IZipReader? zipReader = null)
    {
        _xmlReader = xmlReader ?? new XmlReaderHelpers();
        _zipReader = zipReader ?? new InternalZipReader();
    }

    /// <summary>
    /// Asynchronously opens an Excel file for processing.
    /// </summary>
    /// <param name="fileName">The full path to the Excel file to be opened.</param>
    /// <param name="fileType">The type of the file to be opened. Defaults to <see cref="FileType.Xlsx"/>.</param>
    /// <param name="options">Optional parameters for configuring the file opening process.</param>
    /// <param name="ct">A <see cref="CancellationToken"/> to observe while waiting for the task to complete.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="fileName"/> is <c>null</c>.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be accessed or opened.</exception>
    /// <exception cref="InvalidDataException">Thrown when the file is not a valid Excel file.</exception>
    public Task OpenAsync(in string fileName, FileType fileType = FileType.Xlsx, Options? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileName);
        _fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None, 0x8000/*64*1024*/, true);
        return OpenAsync(_fs, fileType, options, ct);
    }

    /// <InheritDoc />
    public async Task OpenAsync(Stream fileStream, FileType fileType = FileType.Xlsx, Options? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(fileStream);
        if (!fileStream.CanSeek)
        {
            throw new EndOfStreamException("'fileStream' _must_ be seekable!");
        }
        _fs = fileStream;
        await _zipReader.OpenArchiveAsync(fileStream, ct).ConfigureAwait(false);
        // Now perform the Getting of the base data
        var workbook = TempFile.MakeThisATempFile("workbook.xml");
        _baseFiles["xl/workbook.xml"] = workbook;
        using (var targetStream = workbook.FileInfo.OpenWrite())
        {
            await _zipReader.CopyToAsync("xl/workbook.xml", targetStream, ct).ConfigureAwait(false);
        }
        await GetSheetNamesAsync(workbook, ct).ConfigureAwait(false);

        var workbook_rels = TempFile.MakeThisATempFile("workbook.xml.rels");
        _baseFiles["xl/_rels/workbook.xml.rels"] = workbook_rels;
        using (var targetStream = workbook.FileInfo.OpenWrite())
        {
            await _zipReader.CopyToAsync("xl/_rels/workbook.xml.rels", targetStream, ct).ConfigureAwait(false);
        }

        await GetSheetRelationsAsync(workbook_rels, ct).ConfigureAwait(false);

        _sheets = sheets.Where(x => sheetRelations.ContainsKey(x.RelationId))
            .Select(x => new { Sheet = x, ZipEntry = _archive.GetEntry($"xl/{sheetRelations[x.RelationId].Target}") ?? throw new XlsxHelperException($"zip entry not found for {x.SheetName}.") })
            .Select(x => new Worksheet(x.Sheet.SheetName, new WorksheetReader(x.ZipEntry!.Open(), _sharedStringLookup)))
            .ToArray();

    }

    private async Task GetSheetRelationsAsync(TempFile workbook_rels, CancellationToken ct)
    {
        using var stream = workbook_rels.FileInfo.OpenRead();
        var sheetNamesWithrId = new Dictionary<string, (string target, string type)>();
        var stack = new Stack<string>();
        using var reader = await _xmlReader.CreateReaderAsync(stream, ct).ConfigureAwait(false);
        while (await reader.ReadAsync(ct).ConfigureAwait(false)
            && !reader.EOF)
        {
            if (reader.IsElement)
            {
                if (reader.Name == "Relationship"
                    && stack.Count == 1
                    && stack.Peek() == "Relationships"
                    )
                {
                    var target = reader.GetAttribute("Target");
                    var id = reader.GetAttribute("Id");
                    var type = reader.GetAttribute("Type");
                    if (!string.IsNullOrWhiteSpace(target)
                        && !string.IsNullOrWhiteSpace(id)
                        && !string.IsNullOrWhiteSpace(type)
                        )
                    {
                        sheetNamesWithrId.Add(id, (target, type));
                    }
                }

                if (!reader.IsEmptyElement)
                {
                    stack.Push(reader.Name);
                }
            }

            if (reader.IsEndElement)
            {
                stack.Pop();
            }
        }

        var sharedStringTarget = sheetNamesWithrId.Values.FirstOrDefault(x => x.type == XmlNameSpaces.SharedStringRelationshipType).target ?? "sharedStrings.xml";
        var sharedStringEntry = _archive.GetEntry($"xl/{sharedStringTarget}");
        _sharedStringLookup = new SharedStringLookup(sharedStringEntry?.Open() ?? new MemoryStream(), (sharedStringEntry?.Length ?? 0) > 20_000_000);
    }

    private async Task GetSheetNamesAsync(TempFile workbook, CancellationToken ct)
    {
        using var stream = workbook.FileInfo.OpenRead();
        var relationshipNamespace = XmlNameSpaces.RelationshipsOpenXmlFormat;
        using var reader = await _xmlReader.CreateReaderAsync(stream, ct).ConfigureAwait(false);
        var sheetNamesWithrId = new Dictionary<string, int>();
        var stack = new Stack<string>();
        while (await reader.ReadAsync(ct).ConfigureAwait(false)
            && !reader.EOF)
        {
            if (reader.IsElement)
            {
                if (reader.Name == "workbook"
                    && stack.Count == 0
                )
                {
                    relationshipNamespace = reader.GetAttribute("xmlns:r") == XmlNameSpaces.RelationshipsOclc
                        ? XmlNameSpaces.RelationshipsOclc
                        : XmlNameSpaces.RelationshipsOpenXmlFormat;
                }
                if (reader.Name == "sheet"
                    && stack.Count == 2
                    && stack.Peek() == "sheets")
                {
                    var sheetname = reader.GetAttribute("name");
                    var rId = reader.GetAttribute("id", relationshipNamespace);
                    if (!string.IsNullOrWhiteSpace(sheetname) && !string.IsNullOrWhiteSpace(rId))
                    {
                        sheetNamesWithrId.Add(sheetname, Convert.ToInt32(rId, CultureInfo.InvariantCulture));
                    }
                }

                if (!reader.IsEmptyElement)
                {
                    stack.Push(reader.Name);
                }
            }

            if (reader.IsEndElement)
            {
                stack.Pop();
            }
        }
        _sheetNamesWithrId = sheetNamesWithrId.AsReadOnly();
    }


    /// <InheritDoc />
    public IEnumerable<string> SheetNames()
    {
        foreach (var name in _sheetNamesWithrId.Keys)
        {
            yield return name;
        }
    }

    /// <InheritDoc />
    public IAsyncEnumerable<object?[]> GetDefinedRangeAsync(in string rangeName, in string? useThisSheetName = null, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    /// <InheritDoc />
    public Task<ISheet?> GetSheetAsync(in string sheetName, CancellationToken ct = default)
    {
        throw new NotImplementedException();
    }

    protected virtual void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                while (_baseFiles.Any())
                {
                    var tf = _baseFiles.Pop();
                    tf.Dispose();
                }
                while (_sheets.Any())
                {
                    foreach ((var _, var sheet) in _sheets)
                    {
                        sheet.Dispose();
                    }
                    _sheets.Clear();
                }
                _fs?.Dispose();
                _fs = null;
            }
            _isDisposed = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~Excel_PRIME()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
