using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Bing.Offices.Streaming;

/// <summary>为 XLSX 共享字符串提供有界内存、随机读取和磁盘索引。</summary>
internal sealed class XlsxSharedStringStore : IDisposable
{
    private const int IndexRecordSize = sizeof(long) + sizeof(int);
    private static readonly Encoding Utf8 = new UTF8Encoding(false, true);

    private readonly ExcelStreamingTempSession _session;
    private readonly long _memoryBudgetBytes;
    private readonly int _maxCellCharacters;
    private List<string> _memoryValues = new();
    private Dictionary<int, CacheEntry> _cache = new();
    private LinkedList<int> _leastRecentlyUsed = new();
    private Stream _dataFile;
    private Stream _indexFile;
    private long _residentBytes;
    private long _dataLength;
    private bool _hasUnflushedWrites;
    private bool _disposed;

    /// <summary>初始化一个共享字符串存储。</summary>
    /// <param name="session">负责底层临时文件和总字节额度的会话。</param>
    /// <param name="memoryBudgetBytes">字符串及查询缓存共用的内存额度。</param>
    /// <param name="maxCellCharacters">单个共享字符串的最大 .NET 字符数。</param>
    internal XlsxSharedStringStore(ExcelStreamingTempSession session, long memoryBudgetBytes,
        int maxCellCharacters)
    {
        _session = session ?? throw new ArgumentNullException(nameof(session));
        if (memoryBudgetBytes < 0)
            throw new ArgumentOutOfRangeException(nameof(memoryBudgetBytes));
        if (maxCellCharacters < 0)
            throw new ArgumentOutOfRangeException(nameof(maxCellCharacters));
        _memoryBudgetBytes = memoryBudgetBytes;
        _maxCellCharacters = maxCellCharacters;
    }

    /// <summary>获取已加入的共享字符串数量。</summary>
    internal int Count { get; private set; }

    /// <summary>获取是否已迁移到磁盘；一旦迁移便保持为 true。</summary>
    internal bool IsSpilled { get; private set; }

    /// <summary>获取当前字符串列表或 LRU 缓存估算占用的字节数。</summary>
    internal long ResidentBytes => _residentBytes;

    /// <summary>添加字符串，并在内存额度不足时迁移到磁盘。</summary>
    /// <param name="value">已解析的共享字符串。</param>
    internal void Add(string value)
    {
        ThrowIfDisposed();
        ValidateValue(value);
        EnsureCountCanGrow();

        var residentCost = GetResidentCost(value);
        if (!IsSpilled && FitsMemory(residentCost))
        {
            _memoryValues.Add(value);
            _residentBytes += residentCost;
            Count++;
            return;
        }

        if (IsSpilled)
            AppendSpilled(value);
        else
            SpillAndAppend(value);
    }

    /// <summary>异步添加字符串；溢出后的文件写入使用真实异步 IO。</summary>
    /// <param name="value">已解析的共享字符串。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal Task AddAsync(string value, CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
        ValidateValue(value);
        EnsureCountCanGrow();
        if (cancellationToken.IsCancellationRequested)
            return Task.FromCanceled(cancellationToken);

        var residentCost = GetResidentCost(value);
        if (!IsSpilled && FitsMemory(residentCost))
        {
            _memoryValues.Add(value);
            _residentBytes += residentCost;
            Count++;
            return Task.CompletedTask;
        }

        return IsSpilled
            ? AppendSpilledAsync(value, cancellationToken)
            : SpillAndAppendAsync(value, cancellationToken);
    }

    /// <summary>按共享字符串索引读取一个值。</summary>
    /// <param name="index">从零开始的共享字符串索引。</param>
    /// <returns>对应的字符串。</returns>
    /// <exception cref="InvalidDataException">索引不在当前字符串表内，或磁盘记录无效。</exception>
    internal string Get(int index)
    {
        ThrowIfDisposed();
        ValidateIndex(index);
        if (!IsSpilled)
            return _memoryValues[index];
        if (TryGetCached(index, out var cached))
            return cached;

        FlushPendingWrites();
        var record = ReadIndexRecord(index);
        var bytes = new byte[record.ByteLength];
        _dataFile.Position = record.Offset;
        ReadExactly(_dataFile, bytes, 0, bytes.Length);
        var value = Decode(bytes);
        AddToCache(index, value);
        return value;
    }

    /// <summary>异步按共享字符串索引读取一个值；磁盘读取使用真实异步 IO。</summary>
    /// <param name="index">从零开始的共享字符串索引。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>对应的字符串。</returns>
    /// <exception cref="InvalidDataException">索引不在当前字符串表内，或磁盘记录无效。</exception>
    internal Task<string> GetAsync(int index, CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
        ValidateIndex(index);
        if (cancellationToken.IsCancellationRequested)
            return Task.FromCanceled<string>(cancellationToken);
        if (!IsSpilled)
            return Task.FromResult(_memoryValues[index]);
        if (TryGetCached(index, out var cached))
            return Task.FromResult(cached);
        return ReadSpilledAsync(index, cancellationToken);
    }

    /// <summary>关闭本对象持有的流并释放所有内存；临时文件由会话删除。</summary>
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _memoryValues = new List<string>();
        _cache = new Dictionary<int, CacheEntry>();
        _leastRecentlyUsed = new LinkedList<int>();
        _residentBytes = 0;

        var errors = new List<Exception>();
        DisposeFile(_indexFile, errors);
        DisposeFile(_dataFile, errors);
        _indexFile = null;
        _dataFile = null;
        if (errors.Count != 0)
            throw new AggregateException("共享字符串存储流关闭失败。", errors);
    }

    private bool FitsMemory(long valueCost) => valueCost <= _memoryBudgetBytes - _residentBytes;

    private void ValidateValue(string value)
    {
        if (value == null)
            throw new ArgumentNullException(nameof(value));
        if (value.Length > _maxCellCharacters)
            throw new InvalidDataException($"共享字符串超过最大字符数: {_maxCellCharacters}");
    }

    private void EnsureCountCanGrow()
    {
        if (Count == int.MaxValue)
            throw new InvalidDataException("共享字符串数量超过可支持的索引范围。");
    }

    private long GetResidentCost(string value) => checked((long)value.Length * sizeof(char) + 64);

    private void SpillAndAppend(string value)
    {
        Stream data = null;
        Stream index = null;
        try
        {
            data = _session.CreateFile();
            index = _session.CreateFile();
            var dataLength = 0L;
            var recordIndex = 0;
            for (var i = 0; i < _memoryValues.Count; i++)
                dataLength += WriteEntry(data, index, dataLength, recordIndex++, _memoryValues[i]);
            dataLength += WriteEntry(data, index, dataLength, recordIndex, value);
            data.Flush();
            index.Flush();

            _dataFile = data;
            _indexFile = index;
            _dataLength = dataLength;
            IsSpilled = true;
            _memoryValues = new List<string>();
            _residentBytes = 0;
            Count++;
        }
        catch (Exception primary)
        {
            CleanupTemporaryFile(data, primary);
            CleanupTemporaryFile(index, primary);
            throw;
        }
    }

    private async Task SpillAndAppendAsync(string value, CancellationToken cancellationToken)
    {
        Stream data = null;
        Stream index = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            data = _session.CreateFile();
            index = _session.CreateFile();
            var dataLength = 0L;
            var recordIndex = 0;
            for (var i = 0; i < _memoryValues.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                dataLength += await WriteEntryAsync(data, index, dataLength,
                    recordIndex++, _memoryValues[i], cancellationToken).ConfigureAwait(false);
            }
            dataLength += await WriteEntryAsync(data, index, dataLength, recordIndex, value,
                cancellationToken).ConfigureAwait(false);
            await data.FlushAsync(cancellationToken).ConfigureAwait(false);
            await index.FlushAsync(cancellationToken).ConfigureAwait(false);

            _dataFile = data;
            _indexFile = index;
            _dataLength = dataLength;
            IsSpilled = true;
            _memoryValues = new List<string>();
            _residentBytes = 0;
            Count++;
        }
        catch (Exception primary)
        {
            CleanupTemporaryFile(data, primary);
            CleanupTemporaryFile(index, primary);
            throw;
        }
    }

    private void AppendSpilled(string value)
    {
        var dataLength = _dataLength;
        var dataFileLength = _dataFile.Length;
        var indexFileLength = _indexFile.Length;
        try
        {
            _dataLength += WriteEntry(_dataFile, _indexFile, dataLength, Count, value);
            _hasUnflushedWrites = true;
            Count++;
        }
        catch (Exception primary)
        {
            _dataLength = dataLength;
            RollbackAppend(_dataFile, dataFileLength, primary);
            RollbackAppend(_indexFile, indexFileLength, primary);
            throw;
        }
    }

    private async Task AppendSpilledAsync(string value, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var dataLength = _dataLength;
        var dataFileLength = _dataFile.Length;
        var indexFileLength = _indexFile.Length;
        try
        {
            _dataLength += await WriteEntryAsync(_dataFile, _indexFile, dataLength,
                Count, value, cancellationToken).ConfigureAwait(false);
            _hasUnflushedWrites = true;
            Count++;
        }
        catch (Exception primary)
        {
            _dataLength = dataLength;
            RollbackAppend(_dataFile, dataFileLength, primary);
            RollbackAppend(_indexFile, indexFileLength, primary);
            throw;
        }
    }

    private static int WriteEntry(Stream data, Stream index, long offset, int recordIndex, string value)
    {
        var bytes = Utf8.GetBytes(value);
        if (data.Position != offset)
            data.Position = offset;
        data.Write(bytes, 0, bytes.Length);
        WriteIndexRecord(index, recordIndex, offset, bytes.Length);
        return bytes.Length;
    }

    private static async Task<int> WriteEntryAsync(Stream data, Stream index, long offset,
        int recordIndex, string value, CancellationToken cancellationToken)
    {
        var bytes = Utf8.GetBytes(value);
        if (data.Position != offset)
            data.Position = offset;
        await data.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
        var record = CreateIndexRecord(offset, bytes.Length);
        var recordPosition = GetIndexRecordPosition(recordIndex);
        if (index.Position != recordPosition)
            index.Position = recordPosition;
        await index.WriteAsync(record, 0, record.Length, cancellationToken).ConfigureAwait(false);
        return bytes.Length;
    }

    private static void WriteIndexRecord(Stream index, int recordIndex, long offset, int byteLength)
    {
        var record = CreateIndexRecord(offset, byteLength);
        var recordPosition = GetIndexRecordPosition(recordIndex);
        if (index.Position != recordPosition)
            index.Position = recordPosition;
        index.Write(record, 0, record.Length);
    }

    private static long GetIndexRecordPosition(int index) => checked((long)index * IndexRecordSize);

    private static byte[] CreateIndexRecord(long offset, int byteLength)
    {
        var record = new byte[IndexRecordSize];
        BinaryPrimitives.WriteInt64LittleEndian(record.AsSpan(0, sizeof(long)), offset);
        BinaryPrimitives.WriteInt32LittleEndian(record.AsSpan(sizeof(long), sizeof(int)), byteLength);
        return record;
    }

    private IndexRecord ReadIndexRecord(int index)
    {
        var recordPosition = GetIndexRecordPosition(index);
        if (recordPosition > _indexFile.Length - IndexRecordSize)
            throw new InvalidDataException("共享字符串索引记录不完整。");
        var bytes = new byte[IndexRecordSize];
        _indexFile.Position = recordPosition;
        ReadExactly(_indexFile, bytes, 0, bytes.Length);

        var offset = BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(0, sizeof(long)));
        var byteLength = BinaryPrimitives.ReadInt32LittleEndian(
            bytes.AsSpan(sizeof(long), sizeof(int)));
        var dataLength = _dataFile.Length;
        var maxUtf8Bytes = (long)_maxCellCharacters * 3;
        if (offset < 0 || byteLength < 0 || byteLength > maxUtf8Bytes
            || offset > dataLength || byteLength > dataLength - offset)
            throw new InvalidDataException("共享字符串索引记录指向无效数据。");
        return new IndexRecord(offset, byteLength);
    }

    private async Task<string> ReadSpilledAsync(int index, CancellationToken cancellationToken)
    {
        await FlushPendingWritesAsync(cancellationToken).ConfigureAwait(false);
        var recordPosition = GetIndexRecordPosition(index);
        if (recordPosition > _indexFile.Length - IndexRecordSize)
            throw new InvalidDataException("共享字符串索引记录不完整。");
        var record = new byte[IndexRecordSize];
        _indexFile.Position = recordPosition;
        await ReadExactlyAsync(_indexFile, record, 0, record.Length, cancellationToken).ConfigureAwait(false);

        var offset = BinaryPrimitives.ReadInt64LittleEndian(record.AsSpan(0, sizeof(long)));
        var byteLength = BinaryPrimitives.ReadInt32LittleEndian(
            record.AsSpan(sizeof(long), sizeof(int)));
        var dataLength = _dataFile.Length;
        var maxUtf8Bytes = (long)_maxCellCharacters * 3;
        if (offset < 0 || byteLength < 0 || byteLength > maxUtf8Bytes
            || offset > dataLength || byteLength > dataLength - offset)
            throw new InvalidDataException("共享字符串索引记录指向无效数据。");

        var bytes = new byte[byteLength];
        _dataFile.Position = offset;
        await ReadExactlyAsync(_dataFile, bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
        var value = Decode(bytes);
        AddToCache(index, value);
        return value;
    }

    private static void ReadExactly(Stream stream, byte[] buffer, int offset, int count)
    {
        var readTotal = 0;
        while (readTotal < count)
        {
            var read = stream.Read(buffer, offset + readTotal, count - readTotal);
            if (read == 0)
                throw new InvalidDataException("共享字符串数据提前结束。");
            readTotal += read;
        }
    }

    private static async Task ReadExactlyAsync(Stream stream, byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        var readTotal = 0;
        while (readTotal < count)
        {
            var read = await stream.ReadAsync(buffer, offset + readTotal, count - readTotal,
                cancellationToken).ConfigureAwait(false);
            if (read == 0)
                throw new InvalidDataException("共享字符串数据提前结束。");
            readTotal += read;
        }
    }

    private string Decode(byte[] bytes)
    {
        try
        {
            var value = Utf8.GetString(bytes);
            if (value.Length > _maxCellCharacters)
                throw new InvalidDataException("共享字符串数据超过最大字符数。");
            return value;
        }
        catch (DecoderFallbackException exception)
        {
            throw new InvalidDataException("共享字符串数据不是有效的 UTF-8。", exception);
        }
    }

    private bool TryGetCached(int index, out string value)
    {
        if (!_cache.TryGetValue(index, out var entry))
        {
            value = null;
            return false;
        }
        _leastRecentlyUsed.Remove(entry.Node);
        _leastRecentlyUsed.AddFirst(entry.Node);
        value = entry.Value;
        return true;
    }

    private void FlushPendingWrites()
    {
        if (!_hasUnflushedWrites)
            return;
        _dataFile.Flush();
        _indexFile.Flush();
        _hasUnflushedWrites = false;
    }

    private async Task FlushPendingWritesAsync(CancellationToken cancellationToken)
    {
        if (!_hasUnflushedWrites)
            return;
        await _dataFile.FlushAsync(cancellationToken).ConfigureAwait(false);
        await _indexFile.FlushAsync(cancellationToken).ConfigureAwait(false);
        _hasUnflushedWrites = false;
    }

    private void AddToCache(int index, string value)
    {
        var cost = GetResidentCost(value);
        if (cost > _memoryBudgetBytes)
            return;
        while (_residentBytes > _memoryBudgetBytes - cost)
        {
            var leastRecentlyUsed = _leastRecentlyUsed.Last;
            if (leastRecentlyUsed == null)
                break;
            var evicted = _cache[leastRecentlyUsed.Value];
            _residentBytes -= evicted.ResidentCost;
            _cache.Remove(leastRecentlyUsed.Value);
            _leastRecentlyUsed.RemoveLast();
        }
        var node = _leastRecentlyUsed.AddFirst(index);
        _cache.Add(index, new CacheEntry(value, cost, node));
        _residentBytes += cost;
    }

    private void ValidateIndex(int index)
    {
        if (index < 0 || index >= Count)
            throw new InvalidDataException($"共享字符串索引超出范围: {index}");
    }

    private static void RollbackAppend(Stream stream, long length, Exception primary)
    {
        StreamingCleanup.Run(() =>
        {
            stream.Position = 0;
            stream.SetLength(length);
            stream.Position = length;
        }, primary);
    }

    private static void CleanupTemporaryFile(Stream stream, Exception primary)
    {
        if (stream == null)
            return;
        StreamingCleanup.Run(() =>
        {
            stream.Position = 0;
            stream.SetLength(0);
            stream.Position = 0;
        }, primary);
        StreamingCleanup.Run(stream.Dispose, primary);
    }

    private static void DisposeFile(Stream stream, ICollection<Exception> errors)
    {
        if (stream == null)
            return;
        try
        {
            stream.Dispose();
        }
        catch (Exception exception)
        {
            errors.Add(exception);
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(XlsxSharedStringStore));
    }

    private readonly struct IndexRecord
    {
        internal IndexRecord(long offset, int byteLength)
        {
            Offset = offset;
            ByteLength = byteLength;
        }

        internal long Offset { get; }
        internal int ByteLength { get; }
    }

    private sealed class CacheEntry
    {
        internal CacheEntry(string value, long residentCost, LinkedListNode<int> node)
        {
            Value = value;
            ResidentCost = residentCost;
            Node = node;
        }

        internal string Value { get; }
        internal long ResidentCost { get; }
        internal LinkedListNode<int> Node { get; }
    }
}
