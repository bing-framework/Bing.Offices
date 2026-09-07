namespace Bing.Offices.Exports;

/// <summary>不拥有调用方目标流、但在写入边界检查取消的输出流包装器。</summary>
internal sealed class NpoiNonDisposingStream : Stream
{
    /// <summary>由调用方拥有且不会被此包装器释放的底层流。</summary>
    private readonly Stream _inner;
    /// <summary>在写入或刷新边界检查的取消令牌。</summary>
    private readonly CancellationToken _cancellationToken;

    /// <summary>创建不会关闭底层流的 NPOI 流包装器。</summary>
    /// <param name="inner">由调用方负责释放的底层流。</param>
    /// <param name="cancellationToken">写入和刷新期间检查的取消令牌。</param>
    public NpoiNonDisposingStream(Stream inner, CancellationToken cancellationToken = default)
    {
        _inner = inner;
        _cancellationToken = cancellationToken;
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
    public override void Flush()
    {
        _cancellationToken.ThrowIfCancellationRequested();
        _inner.Flush();
    }
    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    /// <inheritdoc />
    public override void SetLength(long value) => _inner.SetLength(value);
    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        _inner.Write(buffer, offset, count);
        _cancellationToken.ThrowIfCancellationRequested();
    }
    /// <summary>刷新底层流但不释放调用方拥有的流。</summary>
    /// <param name="disposing">指示释放流程是否由 Dispose 调用触发。</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
            _inner.Flush();
        base.Dispose(disposing);
    }
}
