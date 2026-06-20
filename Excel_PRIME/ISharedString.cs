namespace ExcelPRIME;

/// <summary>
/// Access the shared string retrieval instance
/// </summary>
public interface ISharedString : IDisposable
{
    /// <summary>
    /// Retrieve the 0 indexed reference from the shared strings
    /// </summary>
    /// <param name="xmlIndex"></param>
    /// <returns></returns>
    string? this[int xmlIndex]
    {
        get;
    }
}
