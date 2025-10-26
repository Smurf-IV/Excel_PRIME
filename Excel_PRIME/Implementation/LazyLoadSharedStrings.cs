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
    private static readonly SemaphoreLocker _locker = new SemaphoreLocker();
    private readonly Stream? _stream;
    private readonly XmlReader _reader;
    private readonly List<string> _currentlyLoaded;
    private bool _isDisposed;
    private readonly string _siRef;
    private readonly string _tRef;
    private readonly StringBuilder _currentStNodeBuilder = new();

    public LazyLoadSharedStrings()
    {
        _currentlyLoaded = new List<string>(0);
        _stream = null;
        _reader = XmlReader.Create(new StringReader(" "), new XmlReaderSettings
        {
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Fragment,
            IgnoreComments = true,
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None
        });
        _siRef = string.Empty;
        _tRef = string.Empty;

    }

    public LazyLoadSharedStrings(Stream stream, CancellationToken ct)
    {
        _stream = stream;
        _reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, // Disable DTDs for untrusted sources
            IgnoreComments = true, // Skip parsing and allocating strings for comments
            IgnoreWhitespace = true, // Ignore significant whitespace
            CheckCharacters = false,
            CloseInput = true,
            ConformanceLevel = ConformanceLevel.Document,
            NameTable = new SharedStringsRestrictedNameTable(),
            ValidationType = ValidationType.None,
            ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
            Async = true // TBD
        });
        // advance to the content
        _reader.ReadToFollowing("sst");

        string? countStr = _reader.GetAttribute("uniqueCount");
        if (!string.IsNullOrEmpty(countStr)
            && int.TryParse(countStr, out int count)
            && count >= 0)
        {
            // Just here to make the logic clearer
        }
        else
        {
            count = 128;
        }

        _currentlyLoaded = new List<string>(count);
        _siRef = _reader.NameTable.Add("si");
        _tRef = _reader.NameTable.Add("t");
    }

    // TODO: Should this be refactored to take a Cancellation Token
    public string? this[int requestIndex]
    {
        get
        {
            if (requestIndex < 0)
            {
                // TODO: Throw an exception ?
                return null;
            }

            // Many sheets may be attempting to get shared strings
            if (requestIndex >= _currentlyLoaded.Count)
            {
                _locker.Lock(() =>
                {
                    // Use additional offset to reduce locking intensity
                    LoadUntil(requestIndex+16);
                    // The "requestIndex >= _currentlyLoaded.Count" is also done internally, so no need to check again after locking
                    if (_reader.EOF
                        || _currentlyLoaded.Count == _currentlyLoaded.Capacity)
                    {
                        // Release resources
                        _reader.Close();
                    }
                });
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

    // TODO: Should this be refactored to take a Cancellation Token
    public string? this[string xmlIndex] => string.IsNullOrEmpty(xmlIndex) ? null : this[xmlIndex.IntParseUnsafe()];

    private void LoadUntil(int untilIndex)
    {
        // TODO: If passed te CancellationToken, should it also be Async ?
        // ReSharper disable once TooWideLocalVariableScope
        string cellValueText;
        while (untilIndex >= _currentlyLoaded.Count
               && _reader.Read()
               && !_reader.EOF
              )
        {
            if (_reader.NodeType == XmlNodeType.Element)
            {
                // Use the pre-atomized string for lightning-fast comparison
                if (Object.ReferenceEquals(_reader.LocalName, _siRef))
                {
                    _currentStNodeBuilder.Clear();
                    int hasMultipleTextForCell = 0;
                    cellValueText = string.Empty;
                    XmlReader subReader = _reader.ReadSubtree();
                    while (subReader.Read()
                           && !subReader.EOF
                          )
                    {
                        if (subReader.NodeType == XmlNodeType.Element)
                        {
                            // Use the pre-atomized string for lightning-fast comparison
                            if (Object.ReferenceEquals(subReader.LocalName, _tRef))
                            {
                                if (subReader.IsEmptyElement)
                                {
                                    continue;
                                }

                                if (hasMultipleTextForCell++ > 0)
                                {
                                    _currentStNodeBuilder.Append(cellValueText);
                                }

                                cellValueText = subReader.ReadElementContentAsString();
                            }
                        }
                    }

                    if (hasMultipleTextForCell > 1)
                    {
                        // Add last iteration, and get current combined string
                        _currentStNodeBuilder.Append(cellValueText);
                        cellValueText = _currentStNodeBuilder.ToString();
                    }

                    _currentlyLoaded.Add(cellValueText);
                }
            }
        }
    }

    private void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                _reader.Dispose();
                _stream?.Dispose();
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
