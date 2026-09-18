using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

/// <summary>
/// 表示 CSV 表头无法与当前映射计划匹配。
/// </summary>
internal sealed class CsvInvalidHeaderException : InvalidOperationException
{
    /// <summary>
    /// 初始化一个 <see cref="CsvInvalidHeaderException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述无效表头的消息。</param>
    public CsvInvalidHeaderException(string message) : base(message) { }
}

/// <summary>
/// 表示 CSV 输入超出配置的资源限制。
/// </summary>
internal sealed class CsvResourceLimitException : InvalidOperationException
{
    /// <summary>
    /// 初始化一个 <see cref="CsvResourceLimitException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述超出资源限制的消息。</param>
    public CsvResourceLimitException(string message) : base(message) { }
}

/// <summary>
/// RFC 4180 风格 CSV 记录读取器。
/// </summary>
internal static class CsvRecordReader
{
    /// <summary>
    /// 按 RFC 4180 规则延迟读取调用方拥有的文本读取器。
    /// </summary>
    /// <param name="reader">提供 CSV 文本的读取器；读取器由调用方负责释放。</param>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="cancellationToken">读取过程中检查的取消令牌。</param>
    /// <returns>按输入顺序产生的 CSV 字段记录。</returns>
    public static IEnumerable<IReadOnlyList<string>> Read(TextReader reader, char delimiter, char quote,
        CancellationToken cancellationToken)
    {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            Mode = CsvMode.RFC4180,
            BadDataFound = _ => throw new InvalidOperationException("CSV 包含不符合 RFC 4180 的字段。")
        };
        using var parser = new CsvParser(reader, configuration, true);
        while (parser.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return parser.Record;
        }
    }

    /// <summary>
    /// 使用 CsvHelper 的异步解析器读取 CSV 记录。
    /// </summary>
    /// <param name="reader">提供 CSV 文本的读取器；读取器由调用方负责释放。</param>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="onRecord">每读取一条记录时异步调用的处理器。</param>
    /// <param name="cancellationToken">读取和回调过程中检查的取消令牌。</param>
    public static async Task ReadAsync(TextReader reader, char delimiter, char quote,
        Func<IReadOnlyList<string>, Task> onRecord, CancellationToken cancellationToken)
    {
        if (reader == null)
            throw new ArgumentNullException(nameof(reader));
        if (onRecord == null)
            throw new ArgumentNullException(nameof(onRecord));
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            Mode = CsvMode.RFC4180,
            BadDataFound = _ => throw new InvalidOperationException("CSV 包含不符合 RFC 4180 的字段。")
        };
        using var parser = new CsvParser(reader, configuration, true);
        while (await parser.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            await onRecord(parser.Record).ConfigureAwait(false);
        }
    }
}

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

/// <summary>
/// 基于 CsvHelper 的 CSV 记录写入器。
/// </summary>
internal sealed class CsvRecordWriter : IDisposable
{
    /// <summary>
    /// 负责 CSV 字段编码和记录格式化的 CsvHelper 写入器。
    /// </summary>
    private readonly CsvWriter _csv;
    /// <summary>
    /// 写入字段时使用的公式注入防护策略。
    /// </summary>
    private readonly CsvFormulaInjectionPolicy _formulaInjectionPolicy;

    /// <summary>
    /// 初始化一个 <see cref="CsvRecordWriter" /> 类型的实例。
    /// </summary>
    /// <param name="writer">接收 CSV 文本的写入器。</param>
    /// <param name="delimiter">字段分隔字符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="newLine">记录换行文本。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    public CsvRecordWriter(TextWriter writer, char delimiter, char quote, string newLine,
        CsvFormulaInjectionPolicy formulaInjectionPolicy)
    {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            NewLine = newLine
        };
        _csv = new CsvWriter(writer, configuration, true);
        _formulaInjectionPolicy = formulaInjectionPolicy;
    }

    /// <summary>
    /// 写入一个字段，并按请求策略保护潜在公式。
    /// </summary>
    /// <param name="field">待写入的字段文本。</param>
    /// <param name="cancellationToken">写入前检查的取消令牌。</param>
    public void WriteField(string field, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _csv.WriteField(ProtectFormula(field ?? string.Empty, _formulaInjectionPolicy));
    }

    /// <summary>
    /// 结束当前记录。
    /// </summary>
    public void NextRecord() => _csv.NextRecord();

    /// <summary>
    /// 异步结束当前记录。
    /// </summary>
    /// <param name="cancellationToken">记录写入过程中检查的取消令牌。</param>
    public async Task NextRecordAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.NextRecordAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 刷新 CsvHelper 和底层文本写入器。
    /// </summary>
    public void Flush() => _csv.Flush();

    /// <summary>
    /// 异步刷新 CsvHelper 和底层文本写入器。
    /// </summary>
    /// <param name="cancellationToken">刷新过程中检查的取消令牌。</param>
    public async Task FlushAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.FlushAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <inheritdoc />
    public void Dispose() => _csv.Dispose();

    /// <summary>
    /// 按公式注入策略保护可能被电子表格解释为公式的字段。
    /// </summary>
    /// <param name="value">待写入 CSV 的字段文本。</param>
    /// <param name="policy">潜在公式字段的处理策略。</param>
    /// <returns>按策略处理后的字段文本。</returns>
    internal static string ProtectFormula(string value, CsvFormulaInjectionPolicy policy)
    {
        if (policy != CsvFormulaInjectionPolicy.Escape || !StartsWithFormula(value))
            return value;
        return $"'{value}";
    }

    /// <summary>
    /// 判断文本首个有效字符是否表示潜在的电子表格公式。
    /// </summary>
    /// <param name="value">待检查的字段文本。</param>
    /// <returns>文本可能被电子表格解释为公式时为 true。</returns>
    private static bool StartsWithFormula(string value)
    {
        var index = 0;
        while (index < value.Length && (value[index] == '\uFEFF' || char.IsWhiteSpace(value[index])
                   || char.IsControl(value[index])))
            index++;
        if (index >= value.Length || "=@+-".IndexOf(value[index]) < 0)
            return false;
        if ((value[index] == '+' || value[index] == '-') && IsSignedNumber(value, index))
            return false;
        return true;
    }

    /// <summary>
    /// 判断带正负号的文本是否为简单十进制数。
    /// </summary>
    /// <param name="value">待检查的字段文本。</param>
    /// <param name="signIndex">正负号在文本中的索引。</param>
    /// <returns>正负号后为完整十进制数字且无其他字符时为 true。</returns>
    private static bool IsSignedNumber(string value, int signIndex)
    {
        var index = signIndex + 1;
        if (index >= value.Length || (value[signIndex] != '+' && value[signIndex] != '-'))
            return false;
        var integerDigits = 0;
        while (index < value.Length && value[index] >= '0' && value[index] <= '9')
        {
            integerDigits++;
            index++;
        }
        var fractionDigits = 0;
        if (index < value.Length && value[index] == '.')
        {
            index++;
            while (index < value.Length && value[index] >= '0' && value[index] <= '9')
            {
                fractionDigits++;
                index++;
            }
        }
        if (integerDigits == 0 && fractionDigits == 0)
            return false;
        if (index < value.Length && (value[index] == 'e' || value[index] == 'E'))
        {
            index++;
            if (index < value.Length && (value[index] == '+' || value[index] == '-'))
                index++;
            var exponentDigits = 0;
            while (index < value.Length && value[index] >= '0' && value[index] <= '9')
            {
                exponentDigits++;
                index++;
            }
            if (exponentDigits == 0)
                return false;
        }
        return index == value.Length;
    }
}

/// <summary>
/// 在内存中完成 CSV 格式化，并通过异步字节写入提交记录。
/// </summary>
internal sealed class CsvAsyncRecordWriter : IDisposable
{
    /// <summary>
    /// 接收编码后 CSV 字节的目标流。
    /// </summary>
    private readonly Stream _destination;
    /// <summary>
    /// 将 CSV 文本转换为目标字节的字符编码。
    /// </summary>
    private readonly Encoding _encoding;
    /// <summary>
    /// 复用的字符编码器，用于跨缓冲区保持编码状态。
    /// </summary>
    private readonly Encoder _encoder;
    /// <summary>
    /// 等待异步刷新到目标流的 CSV 文本缓冲。
    /// </summary>
    private readonly StringBuilder _buffer;
    /// <summary>
    /// 向 CsvHelper 提供文本缓冲的字符串写入器。
    /// </summary>
    private readonly StringWriter _textWriter;
    /// <summary>
    /// 负责 CSV 字段编码和记录格式化的 CsvHelper 写入器。
    /// </summary>
    private readonly CsvWriter _csv;
    /// <summary>
    /// 写入字段时使用的公式注入防护策略。
    /// </summary>
    private readonly CsvFormulaInjectionPolicy _formulaInjectionPolicy;
    /// <summary>
    /// 编码后分块写入目标流的固定大小字节缓冲区。
    /// </summary>
    private readonly byte[] _byteBuffer = new byte[4096];
    /// <summary>
    /// 是否已经向目标流写入编码前导字节标记。
    /// </summary>
    private bool _preambleWritten;
    /// <summary>
    /// 是否已释放此异步记录写入器。
    /// </summary>
    private bool _disposed;

    /// <summary>
    /// 初始化一个 <see cref="CsvAsyncRecordWriter" /> 类型的实例。
    /// </summary>
    /// <param name="destination">接收编码后 CSV 字节的目标流。</param>
    /// <param name="encoding">将 CSV 文本转换为字节的字符编码。</param>
    /// <param name="delimiter">字段分隔字符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="newLine">记录换行文本。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    public CsvAsyncRecordWriter(Stream destination, Encoding encoding, char delimiter, char quote, string newLine,
        CsvFormulaInjectionPolicy formulaInjectionPolicy)
    {
        _destination = destination ?? throw new ArgumentNullException(nameof(destination));
        _encoding = encoding ?? throw new ArgumentNullException(nameof(encoding));
        _encoder = encoding.GetEncoder();
        _preambleWritten = destination.CanSeek && destination.Position > 0;
        _buffer = new StringBuilder();
        _textWriter = new StringWriter(_buffer, CultureInfo.InvariantCulture);
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            NewLine = newLine
        };
        _csv = new CsvWriter(_textWriter, configuration, true);
        _formulaInjectionPolicy = formulaInjectionPolicy;
    }

    /// <summary>
    /// 写入一个字段并按请求策略保护潜在公式。
    /// </summary>
    /// <remarks>
    /// 先写入内存中的 CSV 文本缓冲，后续记录或刷新时提交到目标流。
    /// </remarks>
    /// <param name="field">待写入的字段文本。</param>
    /// <param name="cancellationToken">写入前检查的取消令牌。</param>
    public void WriteField(string field, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _csv.WriteField(CsvRecordWriter.ProtectFormula(field ?? string.Empty, _formulaInjectionPolicy));
    }

    /// <summary>
    /// 完成当前记录，并通过目标流的异步写入接口提交。
    /// </summary>
    /// <param name="cancellationToken">记录写入过程中检查的取消令牌。</param>
    public async Task NextRecordAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.NextRecordAsync().ConfigureAwait(false);
        await WriteBufferedTextAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 刷新 CSV 编码器和目标流，整个 IO 路径使用异步接口。
    /// </summary>
    /// <param name="cancellationToken">刷新过程中检查的取消令牌。</param>
    public async Task FlushAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.FlushAsync().ConfigureAwait(false);
        await WriteBufferedTextAsync(cancellationToken).ConfigureAwait(false);
        await FlushEncoderAsync(cancellationToken).ConfigureAwait(false);
        await _destination.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _csv.Dispose();
        _textWriter.Dispose();
    }

    /// <summary>
    /// 将内存中的 CSV 文本缓冲编码并异步写入目标流。
    /// </summary>
    /// <param name="cancellationToken">写入过程中检查的取消令牌。</param>
    private async Task WriteBufferedTextAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await EnsurePreambleAsync(cancellationToken).ConfigureAwait(false);
        if (_buffer.Length == 0)
            return;
        var chars = _buffer.ToString().ToCharArray();
        _buffer.Clear();
        var charIndex = 0;
        while (charIndex < chars.Length)
        {
            _encoder.Convert(chars, charIndex, chars.Length - charIndex, _byteBuffer, 0,
                _byteBuffer.Length, flush: false, out var charsUsed, out var bytesUsed, out _);
            if (bytesUsed > 0)
                await _destination.WriteAsync(_byteBuffer, 0, bytesUsed, cancellationToken).ConfigureAwait(false);
            charIndex += charsUsed;
            if (charsUsed == 0 && bytesUsed == 0)
                throw new InvalidOperationException("CSV 异步编码器未能推进输入。");
        }
    }

    /// <summary>
    /// 刷新编码器剩余字节并异步写入目标流。
    /// </summary>
    /// <param name="cancellationToken">写入过程中检查的取消令牌。</param>
    private async Task FlushEncoderAsync(CancellationToken cancellationToken)
    {
        var empty = Array.Empty<char>();
        var completed = false;
        while (!completed)
        {
            cancellationToken.ThrowIfCancellationRequested();
            _encoder.Convert(empty, 0, 0, _byteBuffer, 0, _byteBuffer.Length, flush: true,
                out _, out var bytesUsed, out completed);
            if (bytesUsed > 0)
                await _destination.WriteAsync(_byteBuffer, 0, bytesUsed, cancellationToken).ConfigureAwait(false);
            if (!completed && bytesUsed == 0)
                throw new InvalidOperationException("CSV 异步编码器未能完成刷新。");
        }
    }

    /// <summary>
    /// 按目标编码需要写入一次前导字节标记。
    /// </summary>
    /// <param name="cancellationToken">写入过程中检查的取消令牌。</param>
    private async Task EnsurePreambleAsync(CancellationToken cancellationToken)
    {
        if (_preambleWritten)
            return;
        var preamble = _encoding.GetPreamble();
        if (preamble.Length > 0)
            await _destination.WriteAsync(preamble, 0, preamble.Length, cancellationToken).ConfigureAwait(false);
        _preambleWritten = true;
    }
}
