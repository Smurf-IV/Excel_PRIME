using System;
using System.Collections.Generic;

namespace ExcelPRIME.RangeBench;

public interface IGetRange : IDisposable
{
    bool LoadFile(string fullPath);

    IEnumerable<IEnumerable<object?>> GetDefinedRange(string definedName, int? localSheetId = null);

    IEnumerable<IEnumerable<object?>> GetRange(string userRange, string sheetName);
}
