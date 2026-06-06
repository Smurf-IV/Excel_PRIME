using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

using ExcelPRIME.FromExternal;

namespace ExcelPRIME.Implementation;

internal sealed class XmlLazyLoadSharedStrings : ISharedString
{
    private static readonly SemaphoreLocker _locker = new();
    private readonly Stream? _stream;
    private readonly XmlReader _reader;
    private readonly List<string> _currentlyLoaded;
    private bool _isDisposed;
    private readonly string _siRefAtom;
    private readonly string _tRefAtom;
    private readonly StringBuilder _currentStNodeBuilder = new();

    public XmlLazyLoadSharedStrings()
    {
        _currentlyLoaded = [];
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
        _siRefAtom = string.Empty;
        _tRefAtom = string.Empty;
    }

    public XmlLazyLoadSharedStrings(Stream stream, CancellationToken ct)
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
        string sstRefAtom = _reader.NameTable.Add("sst");
        _reader.ReadToFollowing(sstRefAtom);

        string uniqueCountRefAtom = _reader.NameTable.Add("uniqueCount");
        string? countStr = _reader.GetAttribute(uniqueCountRefAtom);
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
        _siRefAtom = _reader.NameTable.Add("si");
        _tRefAtom = _reader.NameTable.Add("t");
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
                _locker.Enter();
                try
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
                }
                finally
                {
                    _locker.Exit();
                }
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

    // TODO: If passed the CancellationToken, should it also be Async ?
    private void LoadUntil(int untilIndex)
    {
        // Parse sequentially until we have loaded enough shared strings.
        while (untilIndex >= _currentlyLoaded.Count
               && _reader.Read()
               && !_reader.EOF
              )
        {
            if (_reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            // Use the pre-atomized string for lightning-fast comparison
            if (!ReferenceEquals(_reader.LocalName, _siRefAtom))
            {
                continue;
            }

            _currentStNodeBuilder.Length = 0;

            if (_reader.IsEmptyElement)
            {
                _currentlyLoaded.Add(string.Empty);
                continue;
            }

            int siDepth = _reader.Depth;

            // Iterate through nodes until we exit the <si> element
            while (_reader.Read())
            {
                // If we've reached the end of <si>, break
                if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == siDepth)
                {
                    break;
                }

                // When we encounter a <t> element, collect its textual content without creating a subtree reader
                if (_reader.NodeType == XmlNodeType.Element && ReferenceEquals(_reader.LocalName, _tRefAtom))
                {
                    if (_reader.IsEmptyElement)
                    {
                        // nothing
                        continue;
                    }

                    int tDepth = _reader.Depth;
                    // Move to the first node inside <t>
                    if (!_reader.Read())
                    {
                        break;
                    }

                    // Collect all text-like nodes until we exit the <t> element
                    while (!_reader.EOF)
                    {
                        if (_reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace)
                        {
                            _currentStNodeBuilder.Append(_reader.Value);
                        }

                        if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == tDepth)
                        {
                            break;
                        }

                        if (!_reader.Read())
                        {
                            break;
                        }
                    }

                    // after inner loop, reader is positioned on the EndElement of <t> (or after), continue outer loop
                }
            }

            _currentlyLoaded.Add(_currentStNodeBuilder.ToString());
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

    public void Dispose() => Dispose(isDisposing: true);
}