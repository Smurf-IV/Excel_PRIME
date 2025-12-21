using System;
using System.Buffers;
using System.Diagnostics;
using System.Text;

namespace ExcelPRIME.XlsbImp;

[DebuggerDisplay("RecordType {RecordType}")]
internal sealed class PooledRecordBuffer : IDisposable
{
    private readonly byte[] _array;
    private bool _isDisposed;
    
    public PooledRecordBuffer(RecordTypeIdentifier recordType, byte[]? array = null, bool succeeded = false)
    {
        RecordType = recordType;
        _array = array!;
        Succeeded = succeeded;
    }

    public ref readonly byte this[int index] => ref _array[index];

    public RecordTypeIdentifier RecordType { get; }

    public bool Succeeded { get; }

    // Return the array to the ArrayPool
    public void Dispose()
    {
        if (!_isDisposed)
        {
            _isDisposed = true;
            if (_array != null!)
            {
                ArrayPool<byte>.Shared.Return(_array);
            }
        }
    }

    public int GetInt32(int offset) => BitConverter.ToInt32(_array, offset);

    public double GetDouble(int offset) => BitConverter.ToDouble(_array, offset);

    public short GetInt16(int offset) => BitConverter.ToInt16(_array, offset);

    public byte GetByte(int offset) => _array[offset];

    public string GetString(int offset)
    {
        int len = BitConverter.ToInt32(_array, offset);
        return Encoding.Unicode.GetString(_array, offset + 4, len * 2);
    }

    public string? GetString(int offset, out int end)
    {
        int len = BitConverter.ToInt32(_array, offset);
        if (len == -1)
        {
            end = offset + 4;
            return null;
        }
        end = offset + 4 + len * 2;
        return Encoding.Unicode.GetString(_array, offset + 4, len * 2);
    }
}