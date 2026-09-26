using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Bing.Offices.Testing.Streams;

/// <summary>
/// 记录流生命周期和实际读写字节数的测试流。
/// </summary>
public class TrackingStream : Stream
{
    /// <summary>
    /// 用于转发读写并跟踪生命周期的底层流。
    /// </summary>
    private readonly Stream _inner;

    /// <summary>
    /// 初始化一个 <see cref="TrackingStream"/> 类型的实例。
    /// </summary>
    /// <param name="inner">需要跟踪读写和生命周期的底层流。</param>
    /// <param name="leaveInnerOpen">释放包装流时是否保留底层流。</param>
    public TrackingStream(Stream inner, bool leaveInnerOpen = false)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        LeaveInnerOpen = leaveInnerOpen;
    }

    /// <summary>
    /// 获取是否保持内部流打开。
    /// </summary>
    public bool LeaveInnerOpen { get; }
    /// <summary>
    /// 获取是否已释放。
    /// </summary>
    public bool WasDisposed { get; private set; }
    /// <summary>
    /// 获取读取字节数。
    /// </summary>
    public long BytesRead { get; private set; }
    /// <summary>
    /// 获取写入字节数。
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
    public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count)
    {
        var read = _inner.Read(buffer, offset, count);
        BytesRead += read;
        return read;
    }
    /// <inheritdoc />
    public override int Read(Span<byte> buffer)
    {
        var read = _inner.Read(buffer);
        BytesRead += read;
        return read;
    }
    /// <inheritdoc />
    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        BytesRead += read;
        return read;
    }
    /// <inheritdoc />
    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
    {
        var read = await _inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
        BytesRead += read;
        return read;
    }
    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    /// <inheritdoc />
    public override void SetLength(long value) => _inner.SetLength(value);
    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count)
    {
        _inner.Write(buffer, offset, count);
        BytesWritten += count;
    }
    /// <inheritdoc />
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        _inner.Write(buffer);
        BytesWritten += buffer.Length;
    }
    /// <inheritdoc />
    public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        await _inner.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
        BytesWritten += count;
    }
    /// <inheritdoc />
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        BytesWritten += buffer.Length;
    }
    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        WasDisposed = true;
        if (disposing && !LeaveInnerOpen)
            _inner.Dispose();
        base.Dispose(disposing);
    }
}

/// <summary>
/// 隐藏底层流定位能力的合同测试包装器。
/// </summary>
public sealed class NonSeekableReadStream : TrackingStream
{
    /// <summary>
    /// 初始化一个 <see cref="NonSeekableReadStream"/> 类型的实例。
    /// </summary>
    /// <param name="inner">需要跟踪读写和生命周期的底层流。</param>
    /// <param name="leaveInnerOpen">释放包装流时是否保留底层流。</param>
    public NonSeekableReadStream(Stream inner, bool leaveInnerOpen = false) : base(inner, leaveInnerOpen) { }
    /// <inheritdoc />
    public override bool CanSeek => false;
    /// <inheritdoc />
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
}
