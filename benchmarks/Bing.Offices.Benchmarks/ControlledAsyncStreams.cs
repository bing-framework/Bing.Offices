namespace Bing.Offices.Benchmarks;

/// <summary>
/// 为异步 IO 基准提供可观测的目标流；每次异步写入都会记录计数和字节数。
/// </summary>
internal abstract class AsyncWriteProbeStream : MemoryStream
{
    public long AsyncWriteCount { get; private set; }

    public long AsyncBytesWritten { get; private set; }

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        => WriteAsyncCore(buffer.AsMemory(offset, count), cancellationToken);

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default)
        => new(WriteAsyncCore(buffer, cancellationToken));

    protected abstract ValueTask WriteChunkAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken);

    private async Task WriteAsyncCore(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await WriteChunkAsync(buffer, cancellationToken).ConfigureAwait(false);
        AsyncWriteCount++;
        AsyncBytesWritten += buffer.Length;
    }
}

/// <summary>在每次异步写入前增加固定延迟的目标流。</summary>
internal sealed class DelayedAsyncWriteStream : AsyncWriteProbeStream
{
    private readonly TimeSpan _delay;
    private readonly int _delayEveryBytes;
    private int _bytesSinceDelay;

    public DelayedAsyncWriteStream(TimeSpan delay, int delayEveryBytes = 64 * 1024)
    {
        if (delay < TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(delay));
        if (delayEveryBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(delayEveryBytes));
        _delay = delay;
        _delayEveryBytes = delayEveryBytes;
    }

    protected override async ValueTask WriteChunkAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken)
    {
        base.Write(buffer.Span);
        _bytesSinceDelay += buffer.Length;
        if (_delay > TimeSpan.Zero && _bytesSinceDelay >= _delayEveryBytes)
        {
            await Task.Delay(_delay, cancellationToken).ConfigureAwait(false);
            _bytesSinceDelay %= _delayEveryBytes;
        }
    }
}

/// <summary>以固定块大小和固定延迟模拟限速异步目标流。</summary>
internal sealed class ThrottledAsyncWriteStream : AsyncWriteProbeStream
{
    private readonly int _chunkBytes;
    private readonly TimeSpan _chunkDelay;
    private int _bytesSinceDelay;

    public ThrottledAsyncWriteStream(int chunkBytes, TimeSpan chunkDelay)
    {
        if (chunkBytes < 1)
            throw new ArgumentOutOfRangeException(nameof(chunkBytes));
        if (chunkDelay < TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(chunkDelay));
        _chunkBytes = chunkBytes;
        _chunkDelay = chunkDelay;
    }

    protected override async ValueTask WriteChunkAsync(ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken)
    {
        while (!buffer.IsEmpty)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var chunkLength = Math.Min(_chunkBytes, buffer.Length);
            base.Write(buffer.Span[..chunkLength]);
            _bytesSinceDelay += chunkLength;
            if (_chunkDelay > TimeSpan.Zero && _bytesSinceDelay >= _chunkBytes)
            {
                await Task.Delay(_chunkDelay, cancellationToken).ConfigureAwait(false);
                _bytesSinceDelay %= _chunkBytes;
            }
            buffer = buffer[chunkLength..];
        }
    }
}
