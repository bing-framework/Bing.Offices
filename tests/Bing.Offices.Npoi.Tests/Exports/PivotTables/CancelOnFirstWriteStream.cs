using System;
using System.IO;
using System.Threading;

namespace Bing.Offices.Tests;

/// <summary>
/// 在首次写入后取消并保留调用方流所有权的测试流。
/// </summary>
internal sealed class CancelOnFirstWriteStream : Stream
{
    /// <summary>
    /// 缓存已转发的 XML 字节以供断言。
    /// </summary>
    private readonly MemoryStream _inner = new();

    /// <summary>
    /// 首次写入后取消的令牌源。
    /// </summary>
    private readonly CancellationTokenSource _cancellation;

    /// <summary>
    /// 初始化一个 <see cref="CancelOnFirstWriteStream" /> 类型的实例。
    /// </summary>
    /// <param name="cancellation">首次写入时触发的取消令牌源。</param>
    internal CancelOnFirstWriteStream(CancellationTokenSource cancellation)
    {
        _cancellation = cancellation ?? throw new ArgumentNullException(nameof(cancellation));
    }

    /// <summary>
    /// 获取已转发到内部流的字节数。
    /// </summary>
    internal long BytesForwarded => _inner.Length;

    /// <summary>
    /// 获取此流是否已被写入器释放。
    /// </summary>
    internal bool WasDisposed { get; private set; }

    /// <inheritdoc />
    public override bool CanRead => false;

    /// <inheritdoc />
    public override bool CanSeek => false;

    /// <inheritdoc />
    public override bool CanWrite => !WasDisposed;

    /// <inheritdoc />
    public override long Length => _inner.Length;

    /// <inheritdoc />
    public override long Position
    {
        get => _inner.Position;
        set => throw new NotSupportedException();
    }

    /// <inheritdoc />
    public override void Flush() => _inner.Flush();

    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    /// <inheritdoc />
    public override void SetLength(long value) => _inner.SetLength(value);

    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count)
    {
        _inner.Write(buffer, offset, count);
        _cancellation.Cancel();
        throw new OperationCanceledException(_cancellation.Token);
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        WasDisposed = true;
        if (disposing)
            _inner.Dispose();
        base.Dispose(disposing);
    }
}
