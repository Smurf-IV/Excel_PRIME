namespace Excel_PRIME;

public record DefinedRange
{
    /// <summary>
    /// Xml.Attribute("name").Value; 
    /// </summary>
    public required string Name { get; init; }
    
    /// <summary>
    /// Xml.Attribute("localSheetId")
    /// </summary>
    public int? SheetIndex { get; set; }
    
    /// <summary>
    /// Xml.Value;
    /// </summary>
    public required string Reference { get; init; }
    
    /// <summary>
    /// Used to generate a key for the dictionary 
    /// </summary>
    public string Key => Name + (!SheetIndex.HasValue ? "" : ":" + SheetIndex);
}
