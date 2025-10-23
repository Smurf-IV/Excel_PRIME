using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal sealed class LazyLoadSharedStrings : ISharedString
{
    private readonly Stream _stream;
    private bool _isDisposed;
    private readonly List<string> _currentlyLoaded;
    private readonly XmlReader _reader;
    private readonly Stack<string> _nodeHierarchy = new();

    public LazyLoadSharedStrings(Stream stream, CancellationToken ct)
    {
        _stream = stream;
        _reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            IgnoreComments = true,
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });
        // advance to the content
        while (_reader.Read())
        {
            if (_reader is { IsEmptyElement: false, NodeType: XmlNodeType.Element })
            {
                _nodeHierarchy.Push(_reader.Name);
            }
            if (ct.IsCancellationRequested
                || _reader is { NodeType: XmlNodeType.Element, LocalName: "sst" })
            {
                break;
            }
        }

        string? countStr = _reader.GetAttribute("uniqueCount");
        if (!string.IsNullOrEmpty(countStr)
            && int.TryParse(countStr, out int count)
            && count >= 0)
        {
        }
        else
        {
            count = 128;
        }

        _currentlyLoaded = new List<string>(count);
    }

    public string? this[string xmlIndex] // TODO: Should this be refactored to take a Cancellation Token
    {
        get
        {
            if (string.IsNullOrEmpty(xmlIndex))
            {
                return null;
            }
            int requestIndex = xmlIndex.IntParseUnsafe();

            if (requestIndex >= _currentlyLoaded.Count)
            {
                LoadUntil(requestIndex);
            }

            if (requestIndex >= _currentlyLoaded.Count)
            {
                // TODO: Throw an exception ?
                return string.Empty;
            }
            else
            {
                return _currentlyLoaded[requestIndex];
            }
        }
    }

    private void LoadUntil(int untilIndex)
    {
        // TODO: If passed te CancellationToken, should it also be Async ?
        StringBuilder currentStNodeBuilder = new();
        // ReSharper disable once TooWideLocalVariableScope
        string cellValueText;
        while (untilIndex >= _currentlyLoaded.Count
               && _reader.ReadToFollowing("si")
               && !_reader.EOF
              )
        {
            currentStNodeBuilder.Clear();
            int hasMultipleTextForCell = 0;
            cellValueText = string.Empty;
            XmlReader subReader = _reader.ReadSubtree();
            while (subReader.ReadToFollowing("t"))
            {
                if (subReader.IsEmptyElement)
                {
                    continue;
                }

                if (hasMultipleTextForCell++ > 0)
                {
                    currentStNodeBuilder.Append(cellValueText);
                }

                cellValueText = subReader.ReadElementContentAsString();
            }
            if (hasMultipleTextForCell > 1)
            {   // Add last iteration, and get current combined string
                currentStNodeBuilder.Append(cellValueText);
                cellValueText = currentStNodeBuilder.ToString();
            }
            _currentlyLoaded.Add(cellValueText);
        }
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _reader.Dispose();
                _stream.Dispose();
            }

            _isDisposed = true;
        }
    }

    ~LazyLoadSharedStrings()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(false);
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }
}
