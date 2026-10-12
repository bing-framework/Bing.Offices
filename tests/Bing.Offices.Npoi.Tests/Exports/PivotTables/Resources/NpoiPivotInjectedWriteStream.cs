using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Resources;

/// <summary>
/// 在目标流接收首批字节后执行注入的失败或取消。
/// </summary>
internal sealed class NpoiPivotInjectedWriteStream : Stream
{
    /// <summary>
    /// 由测试调用方拥有并负责释放的目标流。
    /// </summary>
    private readonly Stream _inner;

    /// <summary>
    /// 首次成功转发写入后执行的测试故障回调。
    /// </summary>
    private readonly Action<int> _afterWrite;

    /// <summary>
    /// 记录首批写入故障是否已经触发。
    /// </summary>
    private bool _triggered;

    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotInjectedWriteStream" /> 类型的实例。
    /// </summary>
    /// <param name="inner">被包装的目标流。</param>
    /// <param name="afterWrite">首次写入后运行的故障回调。</param>
    public NpoiPivotInjectedWriteStream(Stream inner, Action<int> afterWrite)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _afterWrite = afterWrite ?? throw new ArgumentNullException(nameof(afterWrite));
    }

    /// <summary>
    /// 获取已经转发到底层目标流的字节数。
    /// </summary>
    public long BytesWritten { get; private set; }

    /// <summary>
    /// 获取成功转发的非空写入次数。
    /// </summary>
    public int WriteOperations { get; private set; }

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
        AfterWrite(count);
    }

    /// <inheritdoc />
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        _inner.Write(buffer);
        AfterWrite(buffer.Length);
    }

    /// <inheritdoc />
    public override async Task WriteAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        await _inner.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        AfterWrite(count);
    }

    /// <inheritdoc />
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default)
    {
        await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        AfterWrite(buffer.Length);
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        // 底层流由测试调用方拥有；本包装器只拦截写入。
    }

    /// <summary>
    /// 记录写入量并在首次非空写入后触发测试故障。
    /// </summary>
    /// <param name="count">刚转发到底层流的字节数。</param>
    private void AfterWrite(int count)
    {
        BytesWritten += count;
        if (count == 0)
            return;
        WriteOperations++;
        if (_triggered)
            return;
        _triggered = true;
        _afterWrite(count);
    }
}
