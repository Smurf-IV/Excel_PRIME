using System;
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Threading;

using ExcelFAST.Implementation;

using ExcelPRIME;
using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;
using ExcelPRIME.XlsbImp;

namespace ExcelFAST.XlsbImp;

internal sealed class RowFastXlsb : IRowFast
{
    private XlsbStreamReader? _reader;
    private InstanceContext? _instanceContext;
    private (int Height, int Width) _dimensions;
    private int _startRow;
    private bool _isDisposed;

    private Cell[] _cells = Array.Empty<Cell>();
    private int _prevMaxCol;
    private int _bufferedRowOffset = -1;

    public int CurrentRow { get; private set; }

    public void Attach(XlsbSheetReader? sheetReader)
    {
        ClearCells();
        if (sheetReader == null)
        {
            _reader = null;
            _instanceContext = null;
            CurrentRow = 0;
            _bufferedRowOffset = -1;
            return;
        }

        _reader = sheetReader.Reader;
        _instanceContext = sheetReader.InstanceContext;
        _dimensions = sheetReader.SheetDimensions;
        _startRow = sheetReader.StartRow;
        CurrentRow = 0;
        _bufferedRowOffset = -1;
        _prevMaxCol = 0;

        int requiredWidth = Math.Max(_dimensions.Width, 16);
        EnsureCapacity(requiredWidth);
    }

    public void Dispose()
    {
        if (_isDisposed)
        {
            return;
        }

        _isDisposed = true;
        _reader = null;
        _instanceContext = null;
        if (_cells.Length > 0)
        {
            ArrayPool<Cell>.Shared.Return(_cells);
            _cells = Array.Empty<Cell>();
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void ClearCells()
    {
        if (_prevMaxCol > 0 && _cells.Length > 0)
        {
            _cells.AsSpan(0, _prevMaxCol).Clear();
            _prevMaxCol = 0;
        }
    }

    private void EnsureCapacity(int required)
    {
        if (_cells.Length < required)
        {
            int newSize = Math.Max(required, Math.Max(16, _cells.Length * 2));
            Cell[] newArr = ArrayPool<Cell>.Shared.Rent(newSize);
            if (_cells.Length > 0)
            {
                if (_prevMaxCol > 0)
                {
                    _cells.AsSpan(0, _prevMaxCol).CopyTo(newArr);
                }
                ArrayPool<Cell>.Shared.Return(_cells);
            }
            newArr.AsSpan(_prevMaxCol).Clear();
            _cells = newArr;
        }
    }

    public bool GetNextRowData(CancellationToken ct = default)
    {
        if (_reader == null || _isDisposed)
        {
            return false;
        }

        CurrentRow++;

        if (CurrentRow < _startRow)
        {
            ClearCells();
            return true;
        }

        if (_bufferedRowOffset > 0)
        {
            if (_bufferedRowOffset > CurrentRow)
            {
                ClearCells();
                return true;
            }
            if (_bufferedRowOffset == CurrentRow)
            {
                _bufferedRowOffset = -1;
                ParseRowCells(ct);
                return true;
            }
        }

        PooledRecordBuffer nextRecord = _reader.ReadNextRecord();
        while (nextRecord.Succeeded && !ct.IsCancellationRequested)
        {
            if (nextRecord.RecordType == RecordTypeIdentifier.ROWHDR)
            {
                int rowOffset = nextRecord.GetInt32(0) + 1; // 1-based
                nextRecord.Dispose();
                if (rowOffset > CurrentRow)
                {
                    _bufferedRowOffset = rowOffset;
                    ClearCells();
                    return true;
                }

                CurrentRow = rowOffset;
                ParseRowCells(ct);
                return true;
            }

            if (nextRecord.RecordType == RecordTypeIdentifier.DATAEND)
            {
                nextRecord.Dispose();
                return false;
            }

            nextRecord.Dispose();
            nextRecord = _reader.ReadNextRecord();
        }

        nextRecord.Dispose();
        return false;
    }

    private void ParseRowCells(CancellationToken ct)
    {
        ClearCells();

        int maxCol = 0;
        PooledRecordBuffer nextRecord = _reader!.ReadNextRecord();
        while (nextRecord.Succeeded && !ct.IsCancellationRequested
               && nextRecord.RecordType != RecordTypeIdentifier.DATAEND
               && nextRecord.RecordType != RecordTypeIdentifier.ROWHDR)
        {
            RecordTypeIdentifier recordType = nextRecord.RecordType;
            if (recordType is >= RecordTypeIdentifier.CELLBLANK and <= RecordTypeIdentifier.CELLFMLAERROR)
            {
                Cell cell = Cell.ConstructXlsbCell(nextRecord, _instanceContext!);
                if (!cell.CellValue.IsUnknown)
                {
                    int offset = cell.ExcelColumnOffset - 1;
                    if (offset >= 0)
                    {
                        EnsureCapacity(offset + 1);
                        _cells[offset] = cell;
                        if (offset + 1 > maxCol)
                        {
                            maxCol = offset + 1;
                        }
                    }
                }
            }
            nextRecord.Dispose();
            nextRecord = _reader.ReadNextRecord();
        }

        if (nextRecord.RecordType == RecordTypeIdentifier.ROWHDR)
        {
            _reader.RollBackLastRecord(nextRecord);
        }
        else
        {
            nextRecord.Dispose();
        }

        _prevMaxCol = maxCol;

        if (_instanceContext!.Options.ReturnDBNull)
        {
            int width = Math.Max(_dimensions.Width, maxCol);
            EnsureCapacity(width);
            for (int i = 0; i < width; i++)
            {
                if (_cells[i].CellValue.IsUnknown)
                {
                    _cells[i] = new Cell(CellValue.GetDBNull(0), i + 1, CellType.Unknown);
                }
            }
            _prevMaxCol = width;
        }
    }

    public PooledCellView GetCellAt(int excelColumn, CancellationToken ct = default)
    {
        int offset = excelColumn - 1;
        return offset < 0 || offset >= _cells.Length
            ? new PooledCellView(default)
            : new PooledCellView(_cells[offset]);
    }

    public PooledCellView GetCellAt(ReadOnlySpan<char> excelColumn, CancellationToken ct = default)
    {
        int colIndex = ExcelColumns.ParseColumnOffset(excelColumn);
        return GetCellAt(colIndex, ct);
    }

    public ReadOnlySpan<Cell> GetAllCells(CancellationToken ct = default)
    {
        int width = _dimensions.Width > 0 ? _dimensions.Width : _prevMaxCol;
        if (width <= 0)
        {
            return ReadOnlySpan<Cell>.Empty;
        }
        EnsureCapacity(width);
        return _cells.AsSpan(0, width);
    }

    public ReadOnlySpan<Cell> GetCells(int startExcelColumn, int endExcelColumn, CancellationToken ct = default)
    {
        if (startExcelColumn > endExcelColumn || startExcelColumn < 1)
        {
            return ReadOnlySpan<Cell>.Empty;
        }

        int width = _dimensions.Width > 0 ? _dimensions.Width : _prevMaxCol;
        int offset = startExcelColumn - 1;
        if (offset >= width)
        {
            return ReadOnlySpan<Cell>.Empty;
        }

        int count = Math.Min(endExcelColumn - startExcelColumn + 1, width - offset);
        if (count <= 0)
        {
            return ReadOnlySpan<Cell>.Empty;
        }

        EnsureCapacity(offset + count);
        return _cells.AsSpan(offset, count);
    }

    public ReadOnlySpan<Cell> GetCells(ReadOnlySpan<char> startExcelColumn, ReadOnlySpan<char> endExcelColumn,
        CancellationToken ct = default)
    {
        int start = ExcelColumns.ParseColumnOffset(startExcelColumn);
        int end = ExcelColumns.ParseColumnOffset(endExcelColumn);
        return GetCells(start, end, ct);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void CopyBoxedToArray(object?[] values, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(values);
        CopyBoxedToArray(values.AsSpan(), ct);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void CopyBoxedToArray(Span<object?> values, CancellationToken ct = default)
    {
        if (values.IsEmpty)
        {
            return;
        }

        int width = _dimensions.Width > 0 ? _dimensions.Width : _prevMaxCol;
        if (width <= 0)
        {
            values.Clear();
            return;
        }

        EnsureCapacity(width);
        int minLength = Math.Min(values.Length, width);

        ReadOnlySpan<Cell> cellSpan = _cells.AsSpan(0, minLength);
        for (int ordinal = 0; ordinal < minLength; ++ordinal)
        {
            values[ordinal] = cellSpan[ordinal].CellValue.BoxedValue;
        }

        if (values.Length > minLength)
        {
            values.Slice(minLength).Clear();
        }
    }
}
