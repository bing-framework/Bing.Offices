using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;

/// <summary>
/// 在首次转发暂存写入后触发取消。
/// </summary>
internal sealed class AdvancedPivotCancelingWriteStream : Stream
{
    /// <summary>
    /// 由默认提交器拥有的真实暂存文件流。
    /// </summary>
    private readonly Stream _inner;

    /// <summary>
    /// 负责取消当前文件提交的令牌源。
    /// </summary>
    private readonly CancellationTokenSource _cancellation;

    /// <summary>
    /// 记录是否已经触发过写入取消。
    /// </summary>
    private bool _canceled;

    /// <summary>
    /// 初始化一个 <see cref="AdvancedPivotCancelingWriteStream" /> 类型的实例。
    /// </summary>
    /// <param name="inner">由默认提交器管理生命周期的文件流。</param>
    /// <param name="cancellation">写入后取消的令牌源。</param>
    public AdvancedPivotCancelingWriteStream(Stream inner, CancellationTokenSource cancellation)
    {
        _inner = inner;
        _cancellation = cancellation;
    }

    /// <summary>
    /// 获取已经转发到真实暂存流的字节数。
    /// </summary>
    public long BytesWritten { get; private set; }

    /// <inheritdoc />
    public override bool CanRead => _inner.CanRead;

    /// <inheritdoc />
    public override bool CanSeek => _inner.CanSeek;

    /// <inheritdoc />
    public override bool CanWrite => _inner.CanWrite;

    /// <inheritdoc />
    public override long Length => _inner.Length;

    /// <inheritdoc />
    public override long Position { get => _inner.Position; set => _inner.Position = value; }

    /// <inheritdoc />
    public override void Flush() => _inner.Flush();

    /// <inheritdoc />
    public override Task FlushAsync(CancellationToken cancellationToken) =>
        _inner.FlushAsync(cancellationToken);

    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

    /// <inheritdoc />
    public override void SetLength(long value) => _inner.SetLength(value);

    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count)
    {
        _inner.Write(buffer, offset, count);
        CancelAfterWrite(count);
    }

    /// <inheritdoc />
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        _inner.Write(buffer);
        CancelAfterWrite(buffer.Length);
    }

    /// <inheritdoc />
    public override async Task WriteAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        await _inner.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        CancelAfterWrite(count);
    }

    /// <inheritdoc />
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default)
    {
        await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        CancelAfterWrite(buffer.Length);
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        // 默认提交器拥有并释放实际暂存流。
    }

    /// <summary>
    /// 记录真实写入字节并在首次写入后取消。
    /// </summary>
    /// <param name="count">已提交到底层暂存流的字节数。</param>
    private void CancelAfterWrite(int count)
    {
        BytesWritten += count;
        if (_canceled)
            return;
        _canceled = true;
        _cancellation.Cancel();
        throw new OperationCanceledException(_cancellation.Token);
    }
}
