using System;
using System.Buffers;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using ExcelPRIME.Implementation;
using ExcelPRIME.XlsbImp;

using NUnit.Framework;

namespace ExcelPRIME.Tests;

[TestFixture]
internal sealed class XlsbRecordOwnershipTests
{
    [Test]
    public void SyncSheetIterationReleasesSkippedAndRolledBackRecords()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateSheetRecords());
        using BufferedStream stream = new(data);
        /*using*/ XlsbSheetReader reader = new(stream, new InstanceContext(), pool);

        reader.Initialize(CancellationToken.None);
        IRow? first = reader.GetNextRow(RowCellGet.PreGet);
        Assert.That(first, Is.Not.Null);
        Assert.That(first!.GetCell(1).CellValue.AsDouble, Is.EqualTo(12.5));
        first.Dispose();

        IRow? second = reader.GetNextRow(RowCellGet.PreGet);
        Assert.That(second, Is.Not.Null);
        Assert.That(second!.GetCell(2).CellValue.AsDouble, Is.EqualTo(28.25));
        second.Dispose();
        Assert.That(reader.GetNextRow(), Is.Null);

        reader.Dispose();
        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public async Task AsyncSheetIterationReleasesSkippedAndRolledBackRecords()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateSheetRecords());
        using BufferedStream stream = new(data);
        /*using*/ XlsbSheetReader reader = new(stream, new InstanceContext(), pool);

        await reader.InitializeAsync(CancellationToken.None).ConfigureAwait(false);
        IRowAsync? first = await reader.GetNextRowAsync(RowCellGet.PreGet).ConfigureAwait(false);
        Assert.That(first, Is.Not.Null);
        Assert.That(first!.GetCell(1).CellValue.AsDouble, Is.EqualTo(12.5));
        first.Dispose();

        IRowAsync? second = await reader.GetNextRowAsync(RowCellGet.PreGet).ConfigureAwait(false);
        Assert.That(second, Is.Not.Null);
        Assert.That(second!.GetCell(2).CellValue.AsDouble, Is.EqualTo(28.25));
        second.Dispose();
        Assert.That(await reader.GetNextRowAsync().ConfigureAwait(false), Is.Null);

        reader.Dispose();
        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void DisposingSheetReleasesPendingLookaheadRow()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateSheetRecords(firstRowIndex: 1));
        using BufferedStream stream = new(data);
        XlsbSheetReader reader = new(stream, new InstanceContext(), pool);

        reader.Initialize(CancellationToken.None);
        Assert.That(reader.GetNextRow(RowCellGet.PreGet), Is.InstanceOf<INullRow>());
        reader.Dispose();

        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void CellParsingFailureReturnsCurrentRecordBuffer()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateSheetRecords(malformedFirstCell: true));
        using BufferedStream stream = new(data);
        /*using*/ XlsbSheetReader reader = new(stream, new InstanceContext(), pool);
        reader.Initialize(CancellationToken.None);

        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetNextRow(RowCellGet.PreGet));
        reader.Dispose();

        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void SharedStringReadAheadTransfersPendingRecordAndDisposalReleasesIt()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateSharedStringRecords(20));
        XlsbLazyLoadSharedStrings strings = new(data, pool);
        strings.Initialize(CancellationToken.None);

        Assert.That(strings[0], Is.EqualTo("S0"));
        Assert.That(pool.ActiveCount, Is.EqualTo(1));
        Assert.That(strings[17], Is.EqualTo("S17"));
        Assert.That(pool.ActiveCount, Is.Zero);

        strings.Dispose();
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void DisposingSharedStringsReleasesPendingLookaheadRecord()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateSharedStringRecords(20));
        XlsbLazyLoadSharedStrings strings = new(data, pool);
        strings.Initialize(CancellationToken.None);

        Assert.That(strings[0], Is.EqualTo("S0"));
        Assert.That(pool.ActiveCount, Is.EqualTo(1));
        strings.Dispose();

        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void SharedStringReadFailureReturnsCurrentRecordBuffer()
    {
        CountingArrayPool<byte> pool = new();
        using ThrowAtEndStream source = new(CreateSharedStringRecords(1, includeEndRecord: false));
        XlsbLazyLoadSharedStrings strings = new(source, pool);
        strings.Initialize(CancellationToken.None);

        Assert.Throws<IOException>(() => _ = strings[0]);
        Assert.That(pool.ActiveCount, Is.Zero);
        strings.Dispose();
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void MalformedSharedStringReturnsCurrentRecordBuffer()
    {
        CountingArrayPool<byte> pool = new();
        using MemoryStream data = new(CreateMalformedSharedStringRecords());
        XlsbLazyLoadSharedStrings strings = new(data, pool);
        strings.Initialize(CancellationToken.None);

        Assert.Throws<ArgumentOutOfRangeException>(() => _ = strings[0]);
        Assert.That(pool.ActiveCount, Is.Zero);
        strings.Dispose();
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public void SyncReadFailureReturnsRentedRecordBuffer()
    {
        CountingArrayPool<byte> pool = new();
        using ThrowAfterHeaderStream source = new([0x05, 0x08], throwCancellation: false);
        using BufferedStream stream = new(source);
        using XlsbStreamReader reader = new(stream, pool);

        Assert.Throws<IOException>(() => reader.ReadNextRecord());
        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    [Test]
    public async Task AsyncCancellationReturnsRentedRecordBufferAndPropagates()
    {
        CountingArrayPool<byte> pool = new();
        using ThrowAfterHeaderStream source = new([0x05, 0x08], throwCancellation: true);
        using BufferedStream stream = new(source);
        using XlsbStreamReader reader = new(stream, pool);

        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            await reader.ReadNextRecordAsync(CancellationToken.None).ConfigureAwait(false)).ConfigureAwait(false);
        Assert.That(pool.ActiveCount, Is.Zero);
        Assert.That(pool.RentCount, Is.EqualTo(pool.ReturnCount));
    }

    private static byte[] CreateSheetRecords(int firstRowIndex = 0, bool malformedFirstCell = false)
    {
        using MemoryStream stream = new();
        byte[] dimensions = new byte[16];
        BinaryPrimitives.WriteInt32LittleEndian(dimensions, 0);
        BinaryPrimitives.WriteInt32LittleEndian(dimensions.AsSpan(4), firstRowIndex + 1);
        BinaryPrimitives.WriteInt32LittleEndian(dimensions.AsSpan(12), 1);
        WriteRecord(stream, RecordTypeIdentifier.DIMENSION, dimensions);
        WriteRecord(stream, RecordTypeIdentifier.SHEETDATABEGIN, []);
        WriteRecord(stream, RecordTypeIdentifier.ROWHDR, Int32Payload(firstRowIndex));
        WriteRecord(stream, RecordTypeIdentifier.NOTE, [1, 2, 3]);
        WriteRecord(stream, RecordTypeIdentifier.CELLREAL, malformedFirstCell ? Int32Payload(0) : RealCellPayload(0, 12.5));
        WriteRecord(stream, RecordTypeIdentifier.ROWHDR, Int32Payload(firstRowIndex + 1));
        WriteRecord(stream, RecordTypeIdentifier.CELLREAL, RealCellPayload(1, 28.25));
        WriteRecord(stream, RecordTypeIdentifier.DATAEND, []);
        return stream.ToArray();
    }

    private static byte[] CreateSharedStringRecords(int count, bool includeEndRecord = true)
    {
        using MemoryStream stream = new();
        byte[] header = new byte[8];
        BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(4), count);
        WriteRecord(stream, RecordTypeIdentifier.SSTBEGIN, header);
        for (int index = 0; index < count; index++)
        {
            WriteRecord(stream, RecordTypeIdentifier.STRINGITEM, StringItemPayload($"S{index}"));
        }
        if (includeEndRecord)
        {
            WriteRecord(stream, RecordTypeIdentifier.SSTEND, []);
        }
        return stream.ToArray();
    }

    private static byte[] CreateMalformedSharedStringRecords()
    {
        using MemoryStream stream = new();
        byte[] header = new byte[8];
        BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(4), 1);
        WriteRecord(stream, RecordTypeIdentifier.SSTBEGIN, header);
        WriteRecord(stream, RecordTypeIdentifier.STRINGITEM, [0]);
        return stream.ToArray();
    }

    private static byte[] Int32Payload(int value)
    {
        byte[] payload = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(payload, value);
        return payload;
    }

    private static byte[] RealCellPayload(int column, double value)
    {
        byte[] payload = new byte[16];
        BinaryPrimitives.WriteInt32LittleEndian(payload, column);
        BinaryPrimitives.WriteInt64LittleEndian(payload.AsSpan(8), BitConverter.DoubleToInt64Bits(value));
        return payload;
    }

    private static byte[] StringItemPayload(string value)
    {
        byte[] characters = Encoding.Unicode.GetBytes(value);
        byte[] payload = new byte[5 + characters.Length];
        BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(1), value.Length);
        characters.CopyTo(payload, 5);
        return payload;
    }

    private static void WriteRecord(Stream stream, RecordTypeIdentifier recordType, byte[] payload)
    {
        WriteVarInt(stream, (uint)recordType);
        WriteVarInt(stream, (uint)payload.Length);
        stream.Write(payload);
    }

    private static void WriteVarInt(Stream stream, uint value)
    {
        while (value >= 0x80)
        {
            stream.WriteByte((byte)(value | 0x80));
            value >>= 7;
        }
        stream.WriteByte((byte)value);
    }

    private sealed class CountingArrayPool<T> : ArrayPool<T>
    {
        private readonly HashSet<T[]> _active = [];

        public int RentCount { get; private set; }
        public int ReturnCount { get; private set; }
        public int ActiveCount => _active.Count;

        public override T[] Rent(int minimumLength)
        {
            T[] array = new T[minimumLength];
            _active.Add(array);
            RentCount++;
            return array;
        }

        public override void Return(T[] array, bool clearArray = false)
        {
            if (!_active.Remove(array))
            {
                throw new InvalidOperationException("An array was returned more than once or was not rented from this pool.");
            }
            if (clearArray)
            {
                Array.Clear(array);
            }
            ReturnCount++;
        }
    }

    private sealed class ThrowAfterHeaderStream(byte[] header, bool throwCancellation) : MemoryStream(header)
    {
        public override int Read(byte[] buffer, int offset, int count)
        {
            ThrowIfAtPayload();
            return base.Read(buffer, offset, count);
        }

        public override int Read(Span<byte> buffer)
        {
            ThrowIfAtPayload();
            return base.Read(buffer);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            try
            {
                ThrowIfAtPayload();
            }
            catch (Exception exception)
            {
                return ValueTask.FromException<int>(exception);
            }
            return base.ReadAsync(buffer, cancellationToken);
        }

        private void ThrowIfAtPayload()
        {
            if (Position < Length)
            {
                return;
            }
            if (throwCancellation)
            {
                throw new OperationCanceledException();
            }
            throw new IOException("Synthetic XLSB payload read failure.");
        }
    }

    private sealed class ThrowAtEndStream(byte[] content) : MemoryStream(content)
    {
        public override int Read(byte[] buffer, int offset, int count)
        {
            ThrowIfAtEnd();
            return base.Read(buffer, offset, Math.Min(count, 1));
        }

        public override int Read(Span<byte> buffer)
        {
            ThrowIfAtEnd();
            return base.Read(buffer[..Math.Min(buffer.Length, 1)]);
        }

        private void ThrowIfAtEnd()
        {
            if (Position >= Length)
            {
                throw new IOException("Synthetic XLSB end-of-stream read failure.");
            }
        }
    }
}