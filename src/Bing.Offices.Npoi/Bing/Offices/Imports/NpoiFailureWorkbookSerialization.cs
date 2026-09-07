using Bing.Offices.Exceptions;

namespace Bing.Offices.Imports;

/// <summary>Failure Workbook 的序列化、受限输出和取消复制职责。</summary>
internal static class NpoiFailureWorkbookSerialization
{
    /// <summary>将临时工作簿流复制到调用方目标流，并在块边界检查取消。</summary>
    /// <param name="destination">调用方提供的可写目标流。</param>
    /// <param name="source">临时工作簿源流。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    internal static void WriteStream(Stream destination, Stream source, CancellationToken cancellationToken)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
        var buffer = new byte[81920];
        int count;
        while ((count = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
            cancellationToken.ThrowIfCancellationRequested();
        }
    }

    /// <summary>
    /// 查找 NPOI 序列化层包装的大小限制异常。
    /// </summary>
    /// <param name="exception">待检查的异常。</param>
    /// <returns>找到的大小限制异常；否则返回 null。</returns>
    internal static BingOfficesResourceLimitException FindLimitException(Exception exception)
    {
        while (exception != null)
        {
            if (exception is BingOfficesResourceLimitException resourceLimitException)
                return resourceLimitException;
            if (exception is InvalidOperationException invalidOperationException
                && invalidOperationException.Message.StartsWith("失败工作簿超过最大序列化字节数:",
                    StringComparison.Ordinal))
                return new BingOfficesResourceLimitException(invalidOperationException.Message,
                    invalidOperationException, "NPOI", BingOfficesOperation.Import,
                    BingOfficesStage.Serialize);
            exception = exception.InnerException;
        }
        return null;
    }

    /// <summary>查找被 NPOI 包装的取消或致命异常。</summary>
    /// <param name="exception">序列化阶段捕获的异常。</param>
    /// <returns>内部取消或致命异常；不存在时返回 null。</returns>
    internal static Exception FindFatalException(Exception exception)
    {
        while (exception != null)
        {
            if (exception is OperationCanceledException || exception is OutOfMemoryException
                || exception is StackOverflowException)
                return exception;
            exception = exception.InnerException;
        }
        return null;
    }

    /// <summary>
    /// 在序列化阶段限制失败工作簿的最大字节数。
    /// </summary>
    internal sealed class LimitedWriteStream : Stream
    {
        /// <summary>由调用方拥有且仅由包装器刷新、不负责释放的底层输出流。</summary>
        private readonly Stream _inner;
        /// <summary>失败工作簿序列化允许写入的最大字节数。</summary>
        private readonly long? _maxBytes;

        /// <summary>
        /// 初始化受限写入流。
        /// </summary>
        /// <param name="inner">实际写入流。</param>
        /// <param name="maxBytes">最大允许字节数。</param>
        internal LimitedWriteStream(Stream inner, long? maxBytes)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _maxBytes = maxBytes;
        }

        /// <inheritdoc />
        public override bool CanRead => false;

        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;

        /// <inheritdoc />
        public override bool CanWrite => _inner.CanWrite;

        /// <inheritdoc />
        public override long Length => _inner.Length;

        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        /// <inheritdoc />
        public override void SetLength(long value)
        {
            if (_maxBytes.HasValue && value > _maxBytes.Value)
                throw new InvalidOperationException($"失败工作簿超过最大序列化字节数: {_maxBytes.Value}");
            _inner.SetLength(value);
        }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_maxBytes.HasValue && Position > _maxBytes.Value - count)
                throw new InvalidOperationException($"失败工作簿超过最大序列化字节数: {_maxBytes.Value}");
            _inner.Write(buffer, offset, count);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Flush();
            base.Dispose(disposing);
        }
    }
}
