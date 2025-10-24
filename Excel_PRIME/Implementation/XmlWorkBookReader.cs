using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelPRIME.Implementation;

internal sealed class XmlWorkBookReader(XDocument _document) : IXmlWorkBookReader
{
    public void Dispose()
    {
        // TODO release managed resources here
    }

    public Task<IReadOnlyDictionary<string, int>> GetSheetNamesAsync(CancellationToken ct)
    {
        List<XElement> sheetsElements = _document.Descendants()
                    .Where(d => d.Name.LocalName == "sheet")
                    .ToList();
        IReadOnlyDictionary<string, int> dict = sheetsElements.ToDictionary(
            sheetElement => sheetElement.Attributes().First /*OrDefault*/(att => att.Name == "name").Value,
            sheetElement => sheetsElements.IndexOf(sheetElement) + 1
            )
            .AsReadOnly();
        return Task.FromResult(dict);
    }

    public Task<IReadOnlyDictionary<string, DefinedRange>> GetDefinedRangesAsync(CancellationToken ct) => throw new NotImplementedException();
}
