using System;

namespace ExcelPRIME;

public interface ISharedString : IDisposable
{
    string? this[string xmlIndex]
    {
        get;
    }

}
