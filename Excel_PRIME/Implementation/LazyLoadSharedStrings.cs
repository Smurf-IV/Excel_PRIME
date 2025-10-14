using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace ExcelPRIME.Implementation;

internal class LazyLoadSharedStrings : ISharedString
{
    private bool _isDisposed;
    private readonly Dictionary<string, string> _currentlyLoaded = [];
    private readonly XmlReader _reader;
    private readonly Stack<string> _nodeHierarchy = new();

    public LazyLoadSharedStrings(Stream stream, CancellationToken ct)
    {
        _reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            CloseInput = true,
            IgnoreComments = true,
            CheckCharacters = false,
            ConformanceLevel = ConformanceLevel.Document,
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

        var countStr = _reader.GetAttribute("uniqueCount");
        if (!string.IsNullOrEmpty(countStr)
            && int.TryParse(countStr, out int count)
            && count >= 0)
        {
        }
        else
        {
            count = 128;
        }

        _currentlyLoaded = new Dictionary<string, string>(count);
    }

    public string? this[string xmlIndex] // TODO: Should this be refactored to take a Cancellation Token
    {
        get
        {
            if (string.IsNullOrEmpty(xmlIndex))
            {
                return null;
            }
            if (!_currentlyLoaded.TryGetValue(xmlIndex, out var sharedString))
            {
                sharedString = LoadUntil(xmlIndex);
            }
            return sharedString;
        }
    }

    private string? LoadUntil(string xmlIndex)
    {
        // TODO: If passed te CancellationToke, should it also be Async ?
        var untilIndex = Convert.ToInt32(xmlIndex, CultureInfo.InvariantCulture);
        string? lastCellText = null;
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
                var text = _reader.Value;
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
                    var cellText = hasMultipleTextForCell ? currentStNodeBuilder.ToString() : cellValueText;
#pragma warning disable CA1305 // Use the internally faster no culture conversion
                    _currentlyLoaded.Add(_currentlyLoaded.Count.ToString(), cellText!);
#pragma warning restore CA1305
                    lastCellText = cellText;
                    hasMultipleTextForCell = false;
                    cellValueText = null;
                    currentStNodeBuilder.Clear();
                }
            }
        }
        return lastCellText;
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
