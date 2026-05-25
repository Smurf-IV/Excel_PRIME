using ExcelPRIME.Implementation;

namespace ExcelPRIME;

/// <summary>
/// Pass settings etc. down to the classes that need them
/// </summary>
public class InstanceContext
{
    /// <summary>
    /// Access the global shared string retrieval instance
    /// </summary>
    public ISharedString? SharedStrings { get; set; }

    /// <summary>
    /// How to open sheets / convert cells etc.
    /// </summary>
    public Options Options { get; set; } = new Options();

    /// <summary>
    /// What ranges have been defined
    /// </summary>
    public IReadOnlyDictionary<string, DefinedRange>? DefinedRanges { get; set; }

    /// <summary>
    /// Cell styles extracted from the workbook's styles.xml file.
    /// </summary>
    public IReadOnlyDictionary<short, CellStyle>? CellStyles { get; set; }

}
