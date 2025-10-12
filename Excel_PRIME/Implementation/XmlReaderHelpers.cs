using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;


namespace ExcelPRIME.Implementation;

internal sealed class XmlReaderHelpers : IXmlReaderHelpers
{
    /// <InheritDoc />
    public async Task<IReadOnlyList<string>> GetSharedStringsAsync(Stream stream, CancellationToken ct)
    {
        XDocument? document = await XDocument.LoadAsync(stream, LoadOptions.None, ct).ConfigureAwait(false);

        return document != null!
            ? document.Descendants()
                .Where(d => d.Name.LocalName == "si")
                .Select(ReadString)
                .ToList()
            : new List<string>(0).AsReadOnly();
    }

    private static string ReadString(XElement xElement)
    {
        if (!xElement.HasElements)
        {
            return string.Empty;
        }

        return string.Concat(xElement.Descendants()
            .Where(d => d.Name.LocalName == "t")
            .Select(e => XmlConvert.DecodeName(e.Value))
        );
    }

    /// <InheritDoc />
    public async Task<IXmlWorkBookReader?> CreateWorkBookReaderAsync(Stream stream, CancellationToken ct)
    {
        XDocument? document = await XDocument.LoadAsync(stream, LoadOptions.None, ct).ConfigureAwait(false);
        return new XmlWorkBookReader(document);
    }

    /// <InheritDoc />
    public async Task<IXmlSheetReader?> CreateSheetReaderAsync(Stream stream, IReadOnlyList<string> sharedStrings, CancellationToken ct)
    {
        XDocument? document = await XDocument.LoadAsync(stream, LoadOptions.None, ct).ConfigureAwait(false);
        return new XmlSheetReader(document, sharedStrings);
    }

}
