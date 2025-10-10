using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Excel_PRIME;

public interface IXmlWorkBookReader : IDisposable
{
    Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct);
}

public interface IXmlSharedStringsReader : IDisposable
{
}

public interface IXmlSheetReader : IDisposable
{
    /// <summary>
    /// Returns true when the Reader is positioned at the end of the stream. 
    /// </summary>
    bool EOF { get; }

    /// <summary>
    /// Node Properties
    /// Get the type of the current node.
    /// </summary>
    string NodeType { get; }

    /// <summary>
    /// Gets the name of the current node, including the namespace prefix.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Gets the value of the attribute with the specified Name and optional and NamespaceURI
    /// </summary>
    string? GetAttribute(string name, string? namespaceURI = null);

    /// <summary>
    /// Moving through the Stream
    /// Reads the next node from the stream.
    /// </summary>
    Task<bool> ReadAsync(CancellationToken ct);

    /// <summary>
    /// Returns true id this node is an element
    /// </summary>
    bool IsElement { get; }

    /// <summary>
    /// Returns true id this node is an "Empty element"
    /// </summary>
    bool IsEmptyElement { get; }

    /// <summary>
    /// Returns true id this node is an "End element"
    /// </summary>
    bool IsEndElement { get; }


}
