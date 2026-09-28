using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

/// <summary>
/// 对不可定位的 CSV 源流施加读取字节上限且不拥有底层流的包装器。
/// </summary>
internal sealed class CsvLimitedReadStream : Stream
{
    /// <summary>
    /// 调用方拥有的底层输入流。
    /// </summary>
    private readonly Stream _inner;
    /// <summary>
    /// 允许读取的最大字节数。
    /// </summary>
    private readonly long _maxBytes;
    /// <summary>
    /// 已经从底层流读取的字节数。
    /// </summary>
    private long _readBytes;

    /// <summary>
    /// 初始化一个 <see cref="CsvLimitedReadStream" /> 类型的实例。
    /// </summary>
    /// <param name="inner">被限制读取字节数的底层输入流。</param>
    /// <param name="maxBytes">允许读取的最大字节数。</param>
    public CsvLimitedReadStream(Stream inner, long maxBytes)
    {
        _inner = inner;
        _maxBytes = maxBytes;
    }

    /// <inheritdoc />
    public override bool CanRead => _inner.CanRead;
    /// <inheritdoc />
    public override bool CanSeek => false;
    /// <inheritdoc />
    public override bool CanWrite => false;
    /// <inheritdoc />
    public override long Length => throw new NotSupportedException();
    /// <inheritdoc />
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

    /// <inheritdoc />
    /// <remarks>读取总量达到上限后会探测底层流；仍有数据时抛出 <see cref="CsvResourceLimitException" />。</remarks>
    public override int Read(byte[] buffer, int offset, int count)
    {
        if (_readBytes == _maxBytes)
        {
            var probe = _inner.ReadByte();
            if (probe >= 0)
                throw new CsvResourceLimitException($"CSV 输入超过最大字节数: {_maxBytes}");
            return 0;
        }
        var allowed = (int)Math.Min(count, _maxBytes - _readBytes);
        var read = _inner.Read(buffer, offset, allowed);
        _readBytes += read;
        return read;
    }

    /// <inheritdoc />
    /// <remarks>读取总量达到上限后会探测底层流；仍有数据时抛出 <see cref="CsvResourceLimitException" />。</remarks>
    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_readBytes == _maxBytes)
        {
            var probe = await _inner.ReadAsync(new byte[1], 0, 1, cancellationToken).ConfigureAwait(false);
            if (probe > 0)
                throw new CsvResourceLimitException($"CSV 输入超过最大字节数: {_maxBytes}");
            return 0;
        }
        var allowed = (int)Math.Min(count, _maxBytes - _readBytes);
        var read = await _inner.ReadAsync(buffer, offset, allowed, cancellationToken).ConfigureAwait(false);
        _readBytes += read;
        return read;
    }

    /// <inheritdoc />
    public override void Flush() => throw new NotSupportedException();
    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    /// <inheritdoc />
    public override void SetLength(long value) => throw new NotSupportedException();
    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    /// <inheritdoc />
    /// <remarks>包装器不拥有底层流，释放时不会关闭该流。</remarks>
    protected override void Dispose(bool disposing) { }
}
