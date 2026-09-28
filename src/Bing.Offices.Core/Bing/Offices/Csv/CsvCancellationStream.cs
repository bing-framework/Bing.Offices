using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

/// <summary>
/// 将调用方的取消令牌绑定到 CsvHelper/StreamReader 未提供令牌参数的异步入口。
/// </summary>
internal sealed class CsvCancellationStream : Stream
{
    /// <summary>
    /// 被包装且不由此包装器释放的底层流。
    /// </summary>
    private readonly Stream _inner;
    /// <summary>
    /// 异步读写和刷新操作统一使用的取消令牌。
    /// </summary>
    private readonly CancellationToken _cancellationToken;
    /// <summary>
    /// 是否抑制同步刷新，以避免异步路径触发同步 IO。
    /// </summary>
    private readonly bool _suppressSynchronousFlush;

    /// <summary>
    /// 初始化一个 <see cref="CsvCancellationStream" /> 类型的实例。
    /// </summary>
    /// <param name="inner">被包装且由调用方拥有的底层流。</param>
    /// <param name="cancellationToken">异步读写和刷新操作使用的取消令牌。</param>
    /// <param name="suppressSynchronousFlush">是否忽略同步刷新。</param>
    public CsvCancellationStream(Stream inner, CancellationToken cancellationToken,
        bool suppressSynchronousFlush = false)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _cancellationToken = cancellationToken;
        _suppressSynchronousFlush = suppressSynchronousFlush;
    }

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
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

    /// <inheritdoc />
    /// <remarks>使用构造函数绑定的取消令牌进行异步读取，并在读取完成后再次检查取消状态。</remarks>
    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        var read = await _inner.ReadAsync(buffer, offset, count, _cancellationToken).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
        return read;
    }

    /// <inheritdoc />
    /// <remarks>写入前后均检查构造函数绑定的取消令牌。</remarks>
    public override void Write(byte[] buffer, int offset, int count)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        _inner.Write(buffer, offset, count);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    /// <inheritdoc />
    /// <remarks>使用构造函数绑定的取消令牌执行异步写入，并在写入完成后再次检查取消状态。</remarks>
    public override async Task WriteAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        await _inner.WriteAsync(buffer, offset, count, _cancellationToken).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    /// <inheritdoc />
    public override void Flush()
    {
        if (!_suppressSynchronousFlush)
            _inner.Flush();
    }

    /// <inheritdoc />
    /// <remarks>使用构造函数绑定的取消令牌执行异步刷新，并在刷新完成后再次检查取消状态。</remarks>
    public override async Task FlushAsync(CancellationToken cancellationToken)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        await _inner.FlushAsync(_cancellationToken).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    /// <inheritdoc />
    public override void SetLength(long value) => _inner.SetLength(value);

    /// <inheritdoc />
    /// <remarks>包装器不拥有底层流，释放时不会关闭该流。</remarks>
    protected override void Dispose(bool disposing)
    {
        // The caller owns the underlying stream.
        base.Dispose(disposing);
    }
}
