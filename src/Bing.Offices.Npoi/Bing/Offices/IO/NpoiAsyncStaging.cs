using System.IO;

namespace Bing.Offices.IO;

/// <summary>Excel 异步外围 IO 的 staging 策略。</summary>
internal enum NpoiAsyncStagingStrategy
{
    /// <summary>将 staging 内容保存在内存流中。</summary>
    Memory,
    /// <summary>将 staging 内容保存在临时文件中。</summary>
    TempFile,
    /// <summary>以内存为起点，超过阈值后迁移到临时文件。</summary>
    Hybrid
}

/// <summary>可替换的 Excel 异步 staging 工厂，供职责测试和资源探针使用。</summary>
internal interface INpoiAsyncStagingFactory
{
    /// <summary>创建一个新的异步 staging 会话。</summary>
    /// <param name="prefix">临时文件名使用的前缀。</param>
    /// <returns>新建的 staging 会话。</returns>
    INpoiAsyncStaging Create(string prefix);
}

/// <summary>异步 staging 的最小读写合同。</summary>
internal interface INpoiAsyncStaging : IDisposable
{
    /// <summary>获取用于写入 staging 内容的流。</summary>
    Stream WriteStream { get; }

    /// <summary>异步刷新 staging 内容并准备读取。</summary>
    /// <param name="cancellationToken">用于取消刷新的令牌。</param>
    Task FlushAsync(CancellationToken cancellationToken);

    /// <summary>将 staging 内容异步复制到目标流。</summary>
    /// <param name="destination">接收 staging 内容的目标流。</param>
    /// <param name="cancellationToken">用于取消复制的令牌。</param>
    Task CopyToAsync(Stream destination, CancellationToken cancellationToken);
}

/// <summary>
/// 提供 Memory、TempFile 和按阈值迁移的 Hybrid staging 实现。
/// </summary>
internal sealed class NpoiAsyncStagingFactory : INpoiAsyncStagingFactory
{
    /// <summary>初始化一个 <see cref="NpoiAsyncStagingFactory" /> 类型的实例。</summary>
    /// <param name="strategy">内存、临时文件或混合 staging 策略。</param>
    /// <param name="hybridThresholdBytes">混合策略迁移到临时文件的字节阈值。</param>
    /// <param name="directory">临时文件目录；为空时使用系统临时目录。</param>
    internal NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy strategy,
        long hybridThresholdBytes = 8L * 1024 * 1024, string directory = null)
    {
        if (hybridThresholdBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(hybridThresholdBytes));
        Strategy = strategy;
        HybridThresholdBytes = hybridThresholdBytes;
        Directory = directory;
    }

    /// <summary>获取异步暂存策略。</summary>
    internal NpoiAsyncStagingStrategy Strategy { get; }

    /// <summary>获取内存与文件混合暂存阈值（字节）。</summary>
    internal long HybridThresholdBytes { get; }

    /// <summary>获取文件暂存目录。</summary>
    internal string Directory { get; }

    /// <inheritdoc />
    public INpoiAsyncStaging Create(string prefix) => new NpoiAsyncStaging(
        Strategy, HybridThresholdBytes, prefix, Directory);
}

/// <summary>一个 Excel 异步 staging 会话；调用方不拥有会话内部流。</summary>
internal sealed class NpoiAsyncStaging : INpoiAsyncStaging
{
    /// <summary>当前 staging 会话使用的内存、临时文件或混合策略。</summary>
    private readonly NpoiAsyncStagingStrategy _strategy;
    /// <summary>混合策略从内存迁移到临时文件的字节阈值。</summary>
    private readonly long _hybridThresholdBytes;
    /// <summary>创建临时文件时使用的名称前缀。</summary>
    private readonly string _prefix;
    /// <summary>创建临时文件的目标目录；为空时使用系统临时目录。</summary>
    private readonly string _directory;
    /// <summary>内存策略或混合策略迁移前保存暂存内容的内存流。</summary>
    private MemoryStream _memory;
    /// <summary>临时文件策略或混合迁移后承载暂存内容的文件流。</summary>
    private FileStream _writeFile;
    /// <summary>混合策略向调用方公开并负责阈值迁移的写入包装流。</summary>
    private readonly Stream _hybridWriteStream;
    /// <summary>已创建的临时文件路径；尚未创建时为 null。</summary>
    private string _path;
    /// <summary>指示暂存内容是否已完成刷新并准备读取。</summary>
    private bool _flushed;
    /// <summary>指示 staging 会话是否已释放。</summary>
    private bool _disposed;

    /// <summary>初始化一个 <see cref="NpoiAsyncStaging" /> 类型的实例。</summary>
    /// <param name="strategy">内存、临时文件或混合 staging 策略。</param>
    /// <param name="hybridThresholdBytes">混合策略迁移到临时文件的字节阈值。</param>
    /// <param name="prefix">临时文件名使用的前缀。</param>
    /// <param name="directory">临时文件目录；为空时使用系统临时目录。</param>
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

    /// <inheritdoc />
    public Stream WriteStream => _hybridWriteStream ?? _writeFile ?? (Stream)_memory
        ?? throw new ObjectDisposedException(nameof(NpoiAsyncStaging));

    /// <inheritdoc />
    /// <remarks>首次刷新会完成混合暂存迁移并将内存流定位到开头；重复刷新不会重复处理。</remarks>
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

    /// <inheritdoc />
    /// <remarks>复制前会自动刷新暂存内容；目标流由调用方负责释放。</remarks>
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

    /// <inheritdoc />
    /// <remarks>释放内部流并删除已创建的临时文件；清理失败时传播清理异常。</remarks>
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

    /// <summary>
    /// 将内存暂存内容异步迁移到临时文件。
    /// </summary>
    /// <param name="cancellationToken">迁移和刷新过程中检查的取消令牌。</param>
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

    /// <summary>
    /// 将内存暂存内容同步迁移到临时文件。
    /// </summary>
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

    /// <summary>
    /// 将字节数组区间写入混合暂存，并在超过阈值时迁移到临时文件。
    /// </summary>
    /// <param name="buffer">待写入的字节数组。</param>
    /// <param name="offset">源数组中的起始偏移量。</param>
    /// <param name="count">要写入的字节数。</param>
    private void WriteHybrid(byte[] buffer, int offset, int count)
    {
        if (_memory != null && _memory.Position > _hybridThresholdBytes - count)
            MigrateMemoryToTemp();
        (_writeFile ?? (Stream)_memory).Write(buffer, offset, count);
    }

    /// <summary>
    /// 将字节跨度写入混合暂存，并在超过阈值时迁移到临时文件。
    /// </summary>
    /// <param name="buffer">待写入的字节跨度。</param>
    private void WriteHybrid(ReadOnlySpan<byte> buffer)
    {
        if (_memory != null && _memory.Position > _hybridThresholdBytes - buffer.Length)
            MigrateMemoryToTemp();
        (_writeFile ?? (Stream)_memory).Write(buffer);
    }

    /// <summary>
    /// 将单个字节写入混合暂存，并在超过阈值时迁移到临时文件。
    /// </summary>
    /// <param name="value">待写入的字节。</param>
    private void WriteHybridByte(byte value)
    {
        if (_memory != null && _memory.Position >= _hybridThresholdBytes)
            MigrateMemoryToTemp();
        (_writeFile ?? (Stream)_memory).WriteByte(value);
    }

    /// <summary>
    /// 创建并打开异步暂存使用的临时写入流。
    /// </summary>
    private void OpenTempWriteStream()
    {
        var directory = _directory ?? Path.GetTempPath();
        System.IO.Directory.CreateDirectory(directory);
        _path = Path.Combine(directory, $"{_prefix}{Guid.NewGuid():N}.tmp");
        _writeFile = new FileStream(_path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
            81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
    }

    /// <summary>
    /// 检查 staging 会话是否已释放。
    /// </summary>
    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(NpoiAsyncStaging));
    }

    /// <summary>包装 staging 会话，并在写入超过阈值时迁移内容。</summary>
    private sealed class HybridThresholdWriteStream : Stream
    {
        /// <summary>所属 staging 会话，用于转发写入并执行阈值迁移。</summary>
        private readonly NpoiAsyncStaging _owner;

        /// <summary>初始化一个 <see cref="HybridThresholdWriteStream" /> 类型的实例。</summary>
        /// <param name="owner">负责实际 staging 状态和阈值迁移的会话。</param>
        internal HybridThresholdWriteStream(NpoiAsyncStaging owner) => _owner = owner;

        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => true;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _owner._writeFile?.Length ?? _owner._memory.Length;
        /// <inheritdoc />
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

        /// <inheritdoc />
        public override void Flush() => (_owner._writeFile ?? (Stream)_owner._memory).Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) =>
            (_owner._writeFile ?? (Stream)_owner._memory).FlushAsync(cancellationToken);
        /// <inheritdoc />
        /// <remarks>写入操作转发给所属 staging，并在达到阈值时迁移到临时文件。</remarks>
        public override void Write(byte[] buffer, int offset, int count) => _owner.WriteHybrid(buffer, offset, count);
        /// <inheritdoc />
        /// <remarks>写入操作转发给所属 staging，并在达到阈值时迁移到临时文件。</remarks>
        public override void Write(ReadOnlySpan<byte> buffer) => _owner.WriteHybrid(buffer);
        /// <inheritdoc />
        /// <remarks>写入操作转发给所属 staging，并在达到阈值时迁移到临时文件。</remarks>
        public override void WriteByte(byte value) => _owner.WriteHybridByte(value);
        /// <inheritdoc />
        /// <remarks>此写入流不支持读取。</remarks>
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) =>
            (_owner._writeFile ?? (Stream)_owner._memory).Seek(offset, origin);
        /// <inheritdoc />
        /// <remarks>设置长度超过混合阈值时会先迁移到临时文件。</remarks>
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
        /// <inheritdoc />
        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}
