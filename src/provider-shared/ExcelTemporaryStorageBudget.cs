using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Entities;

/// <summary>
/// 为一次实体导入中的受管临时文件提供共享逻辑长度预算。
/// </summary>
internal sealed class ExcelTemporaryStorageBudget
{
    private readonly object _sync = new();
    private readonly long? _maximumBytes;
    private readonly string _provider;
    private long _reservedBytes;

    /// <summary>
    /// 初始化一个临时存储预算。
    /// </summary>
    /// <param name="maximumBytes">最大逻辑字节数；为空时不限制。</param>
    /// <param name="provider">发生限制时报告的 Provider 名称。</param>
    internal ExcelTemporaryStorageBudget(long? maximumBytes, string provider)
    {
        if (maximumBytes.HasValue && maximumBytes.Value <= 0)
            throw new ArgumentOutOfRangeException(nameof(maximumBytes), maximumBytes,
                "临时存储预算必须大于零。");

        _maximumBytes = maximumBytes;
        _provider = string.IsNullOrWhiteSpace(provider) ? "Core" : provider;
    }

    /// <summary>
    /// 创建由该预算跟踪的新临时文件。
    /// </summary>
    /// <param name="path">新临时文件路径。</param>
    /// <param name="options">文件选项。</param>
    /// <param name="bufferSize">文件流缓冲区大小。</param>
    /// <returns>受预算跟踪的流；无预算时返回普通文件流。</returns>
    internal Stream CreateFile(string path, FileOptions options, int bufferSize = 81920)
    {
        if (!_maximumBytes.HasValue)
            return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
                bufferSize, options);

        return new BudgetedFileStream(path, options, bufferSize, this);
    }

    private void Adjust(long oldLength, long newLength)
    {
        if (oldLength == newLength)
            return;

        lock (_sync)
        {
            if (newLength > oldLength)
            {
                var growth = newLength - oldLength;
                if (_reservedBytes > _maximumBytes.Value - growth)
                    throw CreateLimitException();
                _reservedBytes += growth;
                return;
            }

            var reduction = oldLength - newLength;
            if (reduction > _reservedBytes)
                throw new InvalidOperationException("临时存储预算记账不一致。");
            _reservedBytes -= reduction;
        }
    }

    private void Release(long length)
    {
        if (length <= 0)
            return;

        lock (_sync)
        {
            if (length > _reservedBytes)
                throw new InvalidOperationException("临时存储预算释放量超过已预留量。");
            _reservedBytes -= length;
        }
    }

    private BingOfficesResourceLimitException CreateLimitException() =>
        new("实体导入临时文件超过最大临时存储预算。", provider: _provider,
            operation: BingOfficesOperation.Import, stage: BingOfficesStage.Open);

    private sealed class BudgetedFileStream : Stream
    {
        private readonly ExcelTemporaryStorageBudget _budget;
        private readonly FileStream _inner;
        private readonly SemaphoreSlim _gate = new(1, 1);
        private long _accountedLength;
        private int _disposeStarted;
        private int _budgetReleased;

        internal BudgetedFileStream(string path, FileOptions options, int bufferSize,
            ExcelTemporaryStorageBudget budget)
        {
            _budget = budget;
            _inner = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.Read,
                bufferSize, options);
        }

        public override bool CanRead => _inner.CanRead;

        public override bool CanSeek => _inner.CanSeek;

        public override bool CanWrite => _inner.CanWrite;

        public override void Flush()
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                _inner.Flush();
            }
            finally
            {
                _gate.Release();
            }
        }

        public override async Task FlushAsync(CancellationToken cancellationToken)
        {
            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                EnsureOpen();
                await _inner.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (buffer == null)
                throw new ArgumentNullException(nameof(buffer));
            _ = buffer.AsSpan(offset, count);

            _gate.Wait();
            try
            {
                EnsureOpen();
                return _inner.Read(buffer, offset, count);
            }
            finally
            {
                _gate.Release();
            }
        }

        public override int Read(Span<byte> buffer)
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                return _inner.Read(buffer);
            }
            finally
            {
                _gate.Release();
            }
        }

        public override int ReadByte()
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                return _inner.ReadByte();
            }
            finally
            {
                _gate.Release();
            }
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            if (buffer == null)
                throw new ArgumentNullException(nameof(buffer));
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                EnsureOpen();
                return await _inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _gate.Release();
            }
        }

        public override long Length
        {
            get
            {
                _gate.Wait();
                try
                {
                    EnsureOpen();
                    return _inner.Length;
                }
                finally
                {
                    _gate.Release();
                }
            }
        }

        public override long Position
        {
            get
            {
                _gate.Wait();
                try
                {
                    EnsureOpen();
                    return _inner.Position;
                }
                finally
                {
                    _gate.Release();
                }
            }
            set
            {
                _gate.Wait();
                try
                {
                    EnsureOpen();
                    _inner.Position = value;
                }
                finally
                {
                    _gate.Release();
                }
            }
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                return _inner.Seek(offset, origin);
            }
            finally
            {
                _gate.Release();
            }
        }

        public override void SetLength(long value)
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                var oldLength = _accountedLength;
                if (value < oldLength)
                {
                    try
                    {
                        _inner.SetLength(value);
                    }
                    catch
                    {
                        ReconcileShrinkAfterFailure(oldLength);
                        throw;
                    }
                    _budget.Adjust(oldLength, value);
                    _accountedLength = value;
                    return;
                }

                _budget.Adjust(oldLength, value);
                try
                {
                    _inner.SetLength(value);
                    _accountedLength = value;
                }
                catch
                {
                    ReconcileAfterFailure(value);
                    throw;
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        public override void Write(byte[] array, int offset, int count)
        {
            if (array == null)
                throw new ArgumentNullException(nameof(array));
            _ = array.AsSpan(offset, count);

            _gate.Wait();
            try
            {
                EnsureOpen();
                var reservedLength = ReserveWrite(count);
                try
                {
                    _inner.Write(array, offset, count);
                    Reconcile(reservedLength);
                }
                catch
                {
                    ReconcileAfterFailure(reservedLength);
                    throw;
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                var reservedLength = ReserveWrite(buffer.Length);
                try
                {
                    _inner.Write(buffer);
                    Reconcile(reservedLength);
                }
                catch
                {
                    ReconcileAfterFailure(reservedLength);
                    throw;
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        public override void WriteByte(byte value)
        {
            _gate.Wait();
            try
            {
                EnsureOpen();
                var reservedLength = ReserveWrite(1);
                try
                {
                    _inner.WriteByte(value);
                    Reconcile(reservedLength);
                }
                catch
                {
                    ReconcileAfterFailure(reservedLength);
                    throw;
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            if (buffer == null)
                throw new ArgumentNullException(nameof(buffer));
            return WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                EnsureOpen();
                var reservedLength = ReserveWrite(buffer.Length);
                try
                {
                    await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
                    Reconcile(reservedLength);
                }
                catch
                {
                    ReconcileAfterFailure(reservedLength);
                    throw;
                }
            }
            finally
            {
                _gate.Release();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing)
            {
                base.Dispose(false);
                return;
            }

            var shouldDispose = false;
            _gate.Wait();
            try
            {
                shouldDispose = Interlocked.CompareExchange(ref _disposeStarted, 1, 0) == 0;
            }
            finally
            {
                _gate.Release();
            }

            if (!shouldDispose)
            {
                base.Dispose(true);
                return;
            }

            try
            {
                _inner.Dispose();
            }
            finally
            {
                ReleaseBudget();
                base.Dispose(true);
            }
        }

        public override async ValueTask DisposeAsync()
        {
            var shouldDispose = false;
            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                shouldDispose = Interlocked.CompareExchange(ref _disposeStarted, 1, 0) == 0;
            }
            finally
            {
                _gate.Release();
            }

            if (!shouldDispose)
                return;

            try
            {
                await _inner.DisposeAsync().ConfigureAwait(false);
                await base.DisposeAsync().ConfigureAwait(false);
            }
            finally
            {
                ReleaseBudget();
            }
        }

        private long ReserveWrite(int count)
        {
            if (count == 0)
                return _accountedLength;

            var position = _inner.Position;
            if (position < 0 || position > long.MaxValue - count)
                throw _budget.CreateLimitException();

            var targetLength = Math.Max(_accountedLength, position + count);
            _budget.Adjust(_accountedLength, targetLength);
            return targetLength;
        }

        private void Reconcile(long reservedLength)
        {
            long actualLength;
            try
            {
                actualLength = _inner.Length;
            }
            catch
            {
                _accountedLength = reservedLength;
                return;
            }

            if (actualLength != reservedLength)
                _budget.Adjust(reservedLength, actualLength);
            _accountedLength = actualLength;
        }

        private void ReconcileAfterFailure(long reservedLength)
        {
            try
            {
                Reconcile(reservedLength);
            }
            catch
            {
                // 保留完整预留值，避免底层长度无法确认时低估仍存活文件的占用。
                _accountedLength = reservedLength;
            }
        }

        private void ReconcileShrinkAfterFailure(long oldLength)
        {
            try
            {
                var actualLength = _inner.Length;
                if (actualLength < oldLength)
                {
                    _budget.Adjust(oldLength, actualLength);
                    _accountedLength = actualLength;
                }
            }
            catch
            {
                // 保留原预留值，避免清理异常时低估仍存活文件的占用。
            }
        }

        private void EnsureOpen()
        {
            if (Volatile.Read(ref _disposeStarted) != 0)
                throw new ObjectDisposedException(nameof(BudgetedFileStream));
        }

        private void ReleaseBudget()
        {
            if (Interlocked.Exchange(ref _budgetReleased, 1) != 0)
                return;

            _budget.Release(_accountedLength);
            _accountedLength = 0;
        }
    }
}
