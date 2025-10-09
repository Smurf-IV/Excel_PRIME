using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Excel_PRIME.Implementation;

internal class XmlReaderHelpers : IXmlReaderHelpers
{
    public Task<IXmlReader> CreateReaderAsync(Stream stream, CancellationToken ct = default)
    {
        var setting = new XmlReaderSettings
        {
            IgnoreWhitespace = true,
            IgnoreComments = true,
            Async = true
        };
        return Task.Run(() => { IXmlReader internalXmlReader = new InternalXmlReader(XmlReader.Create(stream, setting)); return internalXmlReader; },
            ct);
    }
}
