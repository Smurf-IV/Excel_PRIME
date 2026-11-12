using System;
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Xml;

using ExcelFAST.Implementation;

using ExcelPRIME;
using ExcelPRIME.FromExternal;
using ExcelPRIME.Implementation;

namespace ExcelFAST.XlsxImp;

internal sealed class RowFastXlsx : IRowFast
{
    private XmlReader? _reader;
    private InstanceContext? _instanceContext;
    private ReaderAtoms _readerAtoms;
    private (int Height, int Width) _dimensions;
    private int _startRow;
    private bool _isDisposed;

    private Cell[] _cells = Array.Empty<Cell>();
    private int _prevMaxCol;
    private int _bufferedXmlRowOffset = -1;
    private readonly char[] _charBuffer = new char[256];
    private readonly StringBuilder _valueBuilder = new(256);

    public int CurrentRow { get; private set; }

    public void Attach(XmlSheetReader? sheetReader)
    {
        ClearCells();
        if (sheetReader == null)
        {
            _reader = null;
            _instanceContext = null;
            CurrentRow = 0;
            _bufferedXmlRowOffset = -1;
            return;
        }

        _reader = sheetReader.Reader;
        _instanceContext = sheetReader.InstanceContext;
        _readerAtoms = sheetReader.ReaderAtoms;
        _dimensions = sheetReader.SheetDimensions;
        _startRow = sheetReader.StartRow;
        CurrentRow = 0;
        _bufferedXmlRowOffset = -1;
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
        if (_dimensions.Height > 0 && CurrentRow > _dimensions.Height)
        {
            return false;
        }

        if (CurrentRow < _startRow)
        {
            ClearCells();
            return true;
        }

        if (_bufferedXmlRowOffset > 0)
        {
            if (_bufferedXmlRowOffset > CurrentRow)
            {
                ClearCells();
                return true;
            }
            if (_bufferedXmlRowOffset == CurrentRow)
            {
                _bufferedXmlRowOffset = -1;
                ParseRowCells(ct);
                return true;
            }
        }

        if (!ReadToNextRow(ct))
        {
            return false;
        }

        int rowOffset = ParseRowOffset();
        if (rowOffset > CurrentRow)
        {
            _bufferedXmlRowOffset = rowOffset;
            ClearCells();
            return true;
        }

        CurrentRow = rowOffset;
        ParseRowCells(ct);
        return true;
    }

    private bool ReadToNextRow(CancellationToken ct)
    {
        while (_reader!.ReadToFollowing(_readerAtoms.rowRefAtom) && !ct.IsCancellationRequested)
        {
            if (_reader.NodeType == XmlNodeType.Element && ReferenceEquals(_reader.LocalName, _readerAtoms.rowRefAtom))
            {
                return true;
            }
        }
        return false;
    }

    private int ParseRowOffset()
    {
        int rowOffset = CurrentRow;
        if (_reader!.NodeType == XmlNodeType.Element && ReferenceEquals(_reader.LocalName, _readerAtoms.rowRefAtom))
        {
            while (_reader.MoveToNextAttribute())
            {
                if (ReferenceEquals(_reader.LocalName, _readerAtoms.rRefAtom))
                {
                    rowOffset = _reader.Value.IntParse();
                    break;
                }
            }
            _reader.MoveToElement();
        }
        return rowOffset;
    }

    private void ParseRowCells(CancellationToken ct)
    {
        ClearCells();

        if (_reader!.IsEmptyElement)
        {
            return;
        }

        int currentDepth = _reader.Depth;
        if (_reader.NodeType != XmlNodeType.Element)
        {
            if (_reader.ReadState != ReadState.Interactive)
            {
                return;
            }
            currentDepth--;
        }

        int maxCol = 0;
        while (_reader.Read() && !ct.IsCancellationRequested && _reader.Depth > currentDepth)
        {
            if (_reader.NodeType == XmlNodeType.Element
                && ReferenceEquals(_reader.LocalName, _readerAtoms.cRefAtom)
                && !_reader.IsEmptyElement)
            {
                Cell cell = Cell.ConstructCell(_reader, _instanceContext!, _readerAtoms, _charBuffer, _valueBuilder);
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
}
