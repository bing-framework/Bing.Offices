using System.IO;

namespace Bing.Offices.IO;

/// <summary>Excel 异步外围 IO 的 staging 策略。</summary>
internal enum NpoiAsyncStagingStrategy
{
    Memory,
    TempFile,
    Hybrid
}

/// <summary>可替换的 Excel 异步 staging 工厂，供职责测试和资源探针使用。</summary>
internal interface INpoiAsyncStagingFactory
{
    INpoiAsyncStaging Create(string prefix);
}

/// <summary>异步 staging 的最小读写合同。</summary>
internal interface INpoiAsyncStaging : IDisposable
{
    Stream WriteStream { get; }

    Task FlushAsync(CancellationToken cancellationToken);

    Task CopyToAsync(Stream destination, CancellationToken cancellationToken);
}

/// <summary>
/// 提供 Memory、TempFile 和按阈值迁移的 Hybrid staging 实现。
/// </summary>
internal sealed class NpoiAsyncStagingFactory : INpoiAsyncStagingFactory
{
    internal NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy strategy,
        long hybridThresholdBytes = 8L * 1024 * 1024, string directory = null)
    {
        if (hybridThresholdBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(hybridThresholdBytes));
        Strategy = strategy;
        HybridThresholdBytes = hybridThresholdBytes;
        Directory = directory;
    }

    internal NpoiAsyncStagingStrategy Strategy { get; }

    internal long HybridThresholdBytes { get; }

    internal string Directory { get; }

    public INpoiAsyncStaging Create(string prefix) => new NpoiAsyncStaging(
        Strategy, HybridThresholdBytes, prefix, Directory);
}

/// <summary>一个 Excel 异步 staging 会话；调用方不拥有会话内部流。</summary>
internal sealed class NpoiAsyncStaging : INpoiAsyncStaging
{
    private readonly NpoiAsyncStagingStrategy _strategy;
    private readonly long _hybridThresholdBytes;
    private readonly string _prefix;
    private readonly string _directory;
    private MemoryStream _memory;
    private FileStream _writeFile;
    private readonly Stream _hybridWriteStream;
    private string _path;
    private bool _flushed;
    private bool _disposed;

    internal NpoiAsyncStaging(NpoiAsyncStagingStrategy strategy, long hybridThresholdBytes,
        string prefix, string directory)
    {
        _strategy = strategy;
        _hybridThresholdBytes = hybridThresholdBytes;
        _prefix = string.IsNullOrWhiteSpace(prefix) ? "bing-offices-async-" : prefix;
        _directory = directory;
        if (strategy == NpoiAsyncStagingStrategy.TempFile)
            OpenTempWriteStream();
        else
        {
            _memory = new MemoryStream();
            if (strategy == NpoiAsyncStagingStrategy.Hybrid)
                _hybridWriteStream = new HybridThresholdWriteStream(this);
        }
    }

    public Stream WriteStream => _hybridWriteStream ?? _writeFile ?? (Stream)_memory
        ?? throw new ObjectDisposedException(nameof(NpoiAsyncStaging));

    public async Task FlushAsync(CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
        if (_flushed)
            return;

        if (_strategy == NpoiAsyncStagingStrategy.Hybrid
            && _memory != null && _memory.Length > _hybridThresholdBytes)
            await MigrateMemoryToTempAsync(cancellationToken).ConfigureAwait(false);

        if (_writeFile != null)
        {
            await _writeFile.FlushAsync(cancellationToken).ConfigureAwait(false);
            _writeFile.Dispose();
            _writeFile = null;
        }
        else
        {
            await _memory.FlushAsync(cancellationToken).ConfigureAwait(false);
            _memory.Position = 0;
        }
        _flushed = true;
    }

    public async Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!_flushed)
            await FlushAsync(cancellationToken).ConfigureAwait(false);

        if (_path == null)
        {
            _memory.Position = 0;
            await _memory.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            return;
        }

        using var readFile = new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
        await readFile.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        Exception cleanupException = null;
        try
        {
            _writeFile?.Dispose();
        }
        catch (Exception exception) when (exception is IOException
            || exception is UnauthorizedAccessException)
        {
            cleanupException = exception;
        }
        _writeFile = null;
        try
        {
            _memory?.Dispose();
        }
        catch (Exception exception) when (exception is IOException
            || exception is UnauthorizedAccessException)
        {
            cleanupException ??= exception;
        }
        _memory = null;
        if (_path != null)
        {
            var path = _path;
            _path = null;
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception exception) when (exception is IOException
                || exception is UnauthorizedAccessException)
            {
                if (cleanupException != null)
                    cleanupException.Data["Bing.Offices.AsyncStaging.DeleteException"] = exception;
                else
                    cleanupException = exception;
            }
        }
        if (cleanupException != null)
            throw cleanupException;
    }

    private async Task MigrateMemoryToTempAsync(CancellationToken cancellationToken)
    {
        if (_memory == null)
            return;
        var position = _memory.Position;
        OpenTempWriteStream();
        _memory.Position = 0;
        await _memory.CopyToAsync(_writeFile, 81920, cancellationToken).ConfigureAwait(false);
        _writeFile.Position = position;
        _memory.Dispose();
        _memory = null;
        await _writeFile.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private void MigrateMemoryToTemp()
    {
        if (_memory == null)
            return;
        var position = _memory.Position;
        OpenTempWriteStream();
        _memory.Position = 0;
        _memory.CopyTo(_writeFile);
        _writeFile.Position = position;
        _memory.Dispose();
        _memory = null;
    }

    private void WriteHybrid(byte[] buffer, int offset, int count)
    {
        if (_memory != null && _memory.Position > _hybridThresholdBytes - count)
            MigrateMemoryToTemp();
        (_writeFile ?? (Stream)_memory).Write(buffer, offset, count);
    }

    private void WriteHybrid(ReadOnlySpan<byte> buffer)
    {
        if (_memory != null && _memory.Position > _hybridThresholdBytes - buffer.Length)
            MigrateMemoryToTemp();
        (_writeFile ?? (Stream)_memory).Write(buffer);
    }

    private void WriteHybridByte(byte value)
    {
        if (_memory != null && _memory.Position >= _hybridThresholdBytes)
            MigrateMemoryToTemp();
        (_writeFile ?? (Stream)_memory).WriteByte(value);
    }

    private void OpenTempWriteStream()
    {
        var directory = _directory ?? Path.GetTempPath();
        System.IO.Directory.CreateDirectory(directory);
        _path = Path.Combine(directory, $"{_prefix}{Guid.NewGuid():N}.tmp");
        _writeFile = new FileStream(_path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(NpoiAsyncStaging));
    }

    private sealed class HybridThresholdWriteStream : Stream
    {
        private readonly NpoiAsyncStaging _owner;

        internal HybridThresholdWriteStream(NpoiAsyncStaging owner) => _owner = owner;

        public override bool CanRead => false;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _owner._writeFile?.Length ?? _owner._memory.Length;
        public override long Position
        {
            get => _owner._writeFile?.Position ?? _owner._memory.Position;
            set
            {
                if (_owner._writeFile != null)
                    _owner._writeFile.Position = value;
                else
                    _owner._memory.Position = value;
            }
        }

        public override void Flush() => (_owner._writeFile ?? (Stream)_owner._memory).Flush();
        public override Task FlushAsync(CancellationToken cancellationToken) =>
            (_owner._writeFile ?? (Stream)_owner._memory).FlushAsync(cancellationToken);
        public override void Write(byte[] buffer, int offset, int count) => _owner.WriteHybrid(buffer, offset, count);
        public override void Write(ReadOnlySpan<byte> buffer) => _owner.WriteHybrid(buffer);
        public override void WriteByte(byte value) => _owner.WriteHybridByte(value);
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) =>
            (_owner._writeFile ?? (Stream)_owner._memory).Seek(offset, origin);
        public override void SetLength(long value)
        {
            if (_owner._memory == null)
            {
                _owner._writeFile.SetLength(value);
                return;
            }

            _owner._memory.SetLength(value);
            if (value > _owner._hybridThresholdBytes)
                _owner.MigrateMemoryToTemp();
        }
        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
