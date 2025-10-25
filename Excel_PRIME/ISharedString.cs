using System;

namespace ExcelPRIME;

public interface ISharedString : IDisposable
{
    string? this[int xmlIndex]
    {
        get;
    }

    string? this[string xmlIndex]
    {
        get;
    }

}
