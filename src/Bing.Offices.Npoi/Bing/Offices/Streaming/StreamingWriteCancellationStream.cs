namespace Bing.Offices.Streaming;

/// <summary>为 NPOI 的同步 ZIP 打包写入提供取消检查，不拥有目标流。</summary>
internal sealed class StreamingWriteCancellationStream : Stream
{
    private readonly Stream _inner;
    private readonly CancellationToken _token;
    internal StreamingWriteCancellationStream(Stream inner, CancellationToken token)
    { _inner = inner; _token = token; }
    public override bool CanRead => false;
    public override bool CanSeek => _inner.CanSeek;
    public override bool CanWrite => _inner.CanWrite;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set { _token.ThrowIfCancellationRequested(); _inner.Position = value; } }
    public override void Flush() { _token.ThrowIfCancellationRequested(); _inner.Flush(); }
    public override void Write(byte[] buffer, int offset, int count)
    { _token.ThrowIfCancellationRequested(); _inner.Write(buffer, offset, count); }
    public override void Write(ReadOnlySpan<byte> buffer)
    { _token.ThrowIfCancellationRequested(); _inner.Write(buffer); }
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin)
    { _token.ThrowIfCancellationRequested(); return _inner.Seek(offset, origin); }
    public override void SetLength(long value) { _token.ThrowIfCancellationRequested(); _inner.SetLength(value); }
}
