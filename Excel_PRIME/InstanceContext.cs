using System.Collections.Generic;

namespace ExcelPRIME;

public class InstanceContext
{
    public ISharedString? SharedStrings { get; set; }

    public Options Options { get; set; } = new Options();

    public IReadOnlyDictionary<string, DefinedRange> DefinedRanges { get; set; }

}
