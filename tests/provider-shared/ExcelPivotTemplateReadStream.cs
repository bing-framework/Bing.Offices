using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Bing.Offices.Tests.PivotTemplates;

/// <summary>
/// 在真实模板读取后触发取消或故障。
/// </summary>
internal sealed class ExcelPivotTemplateReadStream : Stream
{
    /// <summary>
    /// 读取后触发的取消源。
    /// </summary>
    private readonly CancellationTokenSource _cancellation;
    /// <summary>
    /// 是否在读取后抛出故障。
    /// </summary>
    private readonly bool _fail;

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateReadStream" /> 类型的实例。
    /// </summary>
    /// <param name="inner">调用方持有的真实文件流。</param>
    /// <param name="cancellation">读取后触发的取消源。</param>
    /// <param name="fail">是否触发读取故障。</param>
    internal ExcelPivotTemplateReadStream(Stream inner, CancellationTokenSource cancellation, bool fail)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _cancellation = cancellation;
        _fail = fail;
    }

    /// <summary>
    /// 获取同步读取次数。
    /// </summary>
    internal int SyncReads { get; private set; }

    /// <summary>
    /// 获取异步读取次数。
    /// </summary>
    internal int AsyncReads { get; private set; }

    /// <summary>
    /// 调用方持有的底层流。
    /// </summary>
    private readonly Stream _inner;

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
    public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count)
    {
        SyncReads++;
        var read = _inner.Read(buffer, offset, count);
        AfterRead();
        return read;
    }
    /// <inheritdoc />
    public override int Read(Span<byte> buffer)
    {
        SyncReads++;
        var read = _inner.Read(buffer);
        AfterRead();
        return read;
    }
    /// <inheritdoc />
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
        ReadAsyncCore(buffer, offset, count, cancellationToken);
    /// <summary>
    /// 执行异步读取并触发测试事件。
    /// </summary>
    private async Task<int> ReadAsyncCore(byte[] buffer, int offset, int count, CancellationToken token)
    {
        AsyncReads++;
        var read = await _inner.ReadAsync(buffer, offset, count, token).ConfigureAwait(false);
        AfterRead();
        return read;
    }

    /// <inheritdoc />
    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        AsyncReads++;
        return ReadMemoryAsync(buffer, cancellationToken);
    }

    /// <summary>
    /// 执行 Memory 异步读取。
    /// </summary>
    private async ValueTask<int> ReadMemoryAsync(Memory<byte> buffer, CancellationToken token)
    {
        var read = await _inner.ReadAsync(buffer, token).ConfigureAwait(false);
        AfterRead();
        return read;
    }

    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    /// <inheritdoc />
    public override void SetLength(long value) => _inner.SetLength(value);
    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
    /// <inheritdoc />
    public override void Write(ReadOnlySpan<byte> buffer) => _inner.Write(buffer);
    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        if (disposing)
            WasDisposed = true;
        base.Dispose(disposing);
    }

    /// <summary>
    /// 获取是否已释放。
    /// </summary>
    internal bool WasDisposed { get; private set; }

    /// <summary>
    /// 触发读取后的测试事件。
    /// </summary>
    private void AfterRead()
    {
        if (_fail)
            throw new IOException("Template snapshot read failed.");
        _cancellation?.Cancel();
    }
}
