using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Excel_PRIME.Implementation;

internal class InternalXmlReader : IXmlReader
{
    private bool _isDisposed;
    private readonly XmlReader _reader;

    public bool EOF => _reader.EOF;

    public string NodeType => _reader.NodeType.ToString();

    public string Name => _reader.Name;

    public bool IsElement => _reader.NodeType == XmlNodeType.Element;

    public bool IsEmptyElement => _reader.IsEmptyElement;

    public bool IsEndElement => _reader.NodeType == XmlNodeType.EndElement;

    public InternalXmlReader(XmlReader reader)
    {
        _reader = reader;
    }

    protected virtual void Dispose(bool isDisposing)
    {
        if (!_isDisposed)
        {
            if (isDisposing)
            {
                // TODO: dispose managed state (managed objects)
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            _isDisposed = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~InternalXmlReader()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(isDisposing: true);
        GC.SuppressFinalize(this);
    }

    public string? GetAttribute(string name, string? namespaceURI = null)
    {
        return string.IsNullOrWhiteSpace(namespaceURI)
            ? _reader.GetAttribute(name)
            : _reader.GetAttribute(name, namespaceURI);
    }

    public Task<bool> ReadAsync(CancellationToken ct) => Task.Run(_reader.ReadAsync, ct);
}
