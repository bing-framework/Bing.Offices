using System.Text;
using System.Xml;

namespace Bing.Offices.Streaming;

/// <summary>同步和异步解析共享的有界 XML 操作。</summary>
internal static class StreamingXml
{
    internal static T Completed<T>(ValueTask<T> operation)
    {
        // 同步分支只取已完成的结果，绝不等待异步 IO。
        if (!operation.IsCompleted) throw new InvalidOperationException("同步 XML 分支产生了未完成操作。");
        return operation.GetAwaiter().GetResult();
    }

    internal static async ValueTask<bool> Read(XmlReader reader, bool asynchronous, int maxDepth, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var read = asynchronous ? await reader.ReadAsync().ConfigureAwait(false) : reader.Read();
        token.ThrowIfCancellationRequested();
        if (reader.Depth > maxDepth) throw new StreamingResourceLimitException("XML 嵌套深度超过限制。");
        return read;
    }

    internal static async ValueTask<string> Text(XmlReader reader, bool asynchronous, int remaining, CancellationToken token)
    {
        var result = new StringBuilder();
        var buffer = new char[1024];
        int count;
        do
        {
            token.ThrowIfCancellationRequested();
            count = asynchronous
                ? await reader.ReadValueChunkAsync(buffer, 0, buffer.Length).ConfigureAwait(false)
                : reader.ReadValueChunk(buffer, 0, buffer.Length);
            if (count > remaining - result.Length) throw new StreamingResourceLimitException("XML 单元格文本超过限制。");
            result.Append(buffer, 0, count);
        } while (count != 0);
        return result.ToString();
    }
}

/// <summary>对实际解压字节计量并将取消传入真实异步读取。</summary>
internal sealed class StreamingReadLimit : Stream
{
    private readonly Stream _inner;
    private readonly long _limit;
    private readonly Action<int> _consume;
    private readonly CancellationToken _token;
    private long _read;
    internal StreamingReadLimit(Stream inner, long limit, CancellationToken token, Action<int> consume = null)
    { _inner = inner; _limit = limit; _token = token; _consume = consume; }
    private int Count(int count)
    {
        if (count > _limit - _read) throw new StreamingResourceLimitException("读取的实际字节数超过限制。");
        _read += count; _consume?.Invoke(count); return count;
    }
    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override int Read(byte[] buffer, int offset, int count)
    { _token.ThrowIfCancellationRequested(); return Count(_inner.Read(buffer, offset, count)); }
    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
    {
        _token.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested();
        return Count(await _inner.ReadAsync(buffer, offset, count, _token).ConfigureAwait(false));
    }
    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
    {
        _token.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested();
        return Count(await _inner.ReadAsync(buffer, _token).ConfigureAwait(false));
    }
    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
}
