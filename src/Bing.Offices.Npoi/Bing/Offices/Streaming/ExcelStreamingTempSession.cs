namespace Bing.Offices.Streaming;

/// <summary>独立调用拥有的临时文件及总字节预算。不管理 NPOI 的系统临时文件。</summary>
internal sealed class ExcelStreamingTempSession : IDisposable
{
    private readonly string _directory;
    private readonly long _limit;
    private readonly List<BudgetFile> _files = new();
    private long _bytes;
    private bool _disposed;

    internal ExcelStreamingTempSession(string parent, long limit)
    {
        _limit = limit;
        _directory = Path.Combine(parent ?? Path.GetTempPath(), "bing-stream-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_directory);
    }

    internal long PeakBytes { get; private set; }

    /// <summary>在复制前为提交器拥有的完整文件副本计费；生命周期仍由提交器负责。</summary>
    internal void ReserveCommitCopy(long bytes) => Charge(bytes);

    internal Stream CreateFile()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(ExcelStreamingTempSession));
        var path = Path.Combine(_directory, _files.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".tmp");
        var file = new BudgetFile(this, path);
        _files.Add(file);
        return file;
    }

    private void Charge(long delta)
    {
        if (delta > _limit - _bytes)
            throw new StreamingResourceLimitException("库自有临时文件总字节数超过限制。");
        _bytes += delta;
        PeakBytes = Math.Max(PeakBytes, _bytes);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        var errors = new List<Exception>();
        foreach (var file in _files)
        {
            try { file.Dispose(); } catch (Exception e) { errors.Add(e); }
            try { File.Delete(file.Path); } catch (Exception e) { errors.Add(e); }
        }
        try { Directory.Delete(_directory); } catch (Exception e) { errors.Add(e); }
        if (errors.Count != 0) throw new AggregateException("流式 Excel 临时文件清理失败。", errors);
    }

    private sealed class BudgetFile : Stream
    {
        private readonly ExcelStreamingTempSession _owner;
        private readonly FileStream _inner;
        private long _reservedLength;
        internal string Path { get; }
        internal BudgetFile(ExcelStreamingTempSession owner, string path)
        {
            _owner = owner; Path = path;
            _inner = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
                65536, FileOptions.Asynchronous);
        }
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override void Flush() => _inner.Flush();
        public override Task FlushAsync(CancellationToken token) => _inner.FlushAsync(token);
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
            => _inner.ReadAsync(buffer, offset, count, token);
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
            => _inner.ReadAsync(buffer, token);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        private void Reserve(int count)
        {
            var required = checked(Position + count);
            if (required <= _reservedLength) return;
            _owner.Charge(required - _reservedLength);
            _reservedLength = required;
        }
        public override void SetLength(long value)
        {
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
            _owner.Charge(value - _reservedLength);
            _reservedLength = value;
            _inner.SetLength(value);
        }
        public override void Write(byte[] buffer, int offset, int count)
        { Reserve(count); _inner.Write(buffer, offset, count); }
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
        { token.ThrowIfCancellationRequested(); Reserve(count); return _inner.WriteAsync(buffer, offset, count, token); }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
        { token.ThrowIfCancellationRequested(); Reserve(buffer.Length); return _inner.WriteAsync(buffer, token); }
        protected override void Dispose(bool disposing)
        { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}

/// <summary>流式处理的明确资源边界。</summary>
internal sealed class StreamingResourceLimitException : InvalidOperationException
{
    internal StreamingResourceLimitException(string message) : base(message) { }
}

/// <summary>清理错误不得覆盖调用方的取消或业务异常。</summary>
internal static class StreamingCleanup
{
    /// <summary>尝试释放全部资源，保留每个清理异常及原始处理异常。</summary>
    internal static void RunAll(Exception primary, params Action[] cleanups)
        => Run(() =>
        {
            var failures = new List<Exception>();
            foreach (var cleanup in cleanups)
                try { cleanup(); } catch (Exception error) { failures.Add(error); }
            if (failures.Count != 0) throw new AggregateException("流式资源清理失败。", failures);
        }, primary);

    internal static void Run(Action cleanup, Exception primary)
    {
        try { cleanup(); }
        catch (Exception error)
        {
            if (primary == null) throw;
            const string key = "Bing.Offices.Streaming.CleanupErrors";
            if (primary.Data[key] is not List<Exception> errors)
                primary.Data[key] = errors = new List<Exception>();
            errors.Add(error);
        }
    }
}
