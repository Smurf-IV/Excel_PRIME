using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

using ExcelPRIME.Shared;

namespace ExcelPRIME.Implementation;

internal class LazyLoadSharedStrings : ISharedString
{
    private bool _isDisposed;
    private readonly List<string> _currentlyLoaded;
    private readonly XmlReader _reader;
    private readonly Stack<string> _nodeHierarchy = new();

    public LazyLoadSharedStrings(Stream stream, CancellationToken ct)
    {
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
            return _currentlyLoaded[requestIndex];
        }
    }

    private void LoadUntil(int untilIndex)
    {
        // TODO: If passed te CancellationToke, should it also be Async ?
        bool hasMultipleTextForCell = false;
        string? cellValueText = null;
        StringBuilder currentStNodeBuilder = new();
        while (untilIndex >= _currentlyLoaded.Count
            && _reader.Read()
            && !_reader.EOF
            )
        {
            if (IsSiTextNode(_nodeHierarchy)
                || IsSiRichTextNode(_nodeHierarchy)
                )
            {
                //there can be multiple `t` tags for each `si` node, in that case combine all.
                string text = _reader.Value;
                if (cellValueText != null)
                {
                    if (!hasMultipleTextForCell)
                    {
                        currentStNodeBuilder.Append(cellValueText);
                    }

                    hasMultipleTextForCell = true;
                    currentStNodeBuilder.Append(text);
                }
                cellValueText = text;
            }

            if (_reader is { IsEmptyElement: false, NodeType: XmlNodeType.Element })
            {
                _nodeHierarchy.Push(_reader.Name);
            }

            if (_reader.NodeType == XmlNodeType.EndElement)
            {
                _nodeHierarchy.Pop();

                if (IsSiElementNode(_nodeHierarchy))
                {
                    string? cellText = hasMultipleTextForCell ? currentStNodeBuilder.ToString() : cellValueText;
                    _currentlyLoaded.Add(cellText!);
                    hasMultipleTextForCell = false;
                    cellValueText = null;
                    currentStNodeBuilder.Clear();
                }
            }
        }
    }

    private bool IsSiElementNode(Stack<string> nodeHierarchy)
        => _reader.Name == "si" && nodeHierarchy.Count == 1 && nodeHierarchy.Peek() == "sst";

    private bool IsSiTextNode(Stack<string> nodeHierarchy)
        => _reader.NodeType is XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace
           && nodeHierarchy.Count == 3 && nodeHierarchy.Peek() == "t";

    private bool IsSiRichTextNode(Stack<string> nodeHierarchy)
        => _reader.NodeType is XmlNodeType.Text or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace
           && nodeHierarchy.Count == 4 && nodeHierarchy.Peek() == "t";

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _reader.Dispose();
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
