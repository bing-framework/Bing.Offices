using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

/// <summary>表示 CSV 表头无法与当前映射计划匹配。</summary>
internal sealed class CsvInvalidHeaderException : InvalidOperationException
{
    /// <summary>使用表头结构错误消息初始化异常。</summary>
    /// <param name="message">描述无效表头的消息。</param>
    public CsvInvalidHeaderException(string message) : base(message) { }
}

/// <summary>表示 CSV 输入超出配置的资源限制。</summary>
internal sealed class CsvResourceLimitException : InvalidOperationException
{
    /// <summary>使用资源限制错误消息初始化异常。</summary>
    /// <param name="message">描述超出资源限制的消息。</param>
    public CsvResourceLimitException(string message) : base(message) { }
}

/// <summary>RFC 4180 风格 CSV 记录读取器。</summary>
internal static class CsvRecordReader
{
    /// <summary>按 RFC 4180 规则延迟读取调用方拥有的文本读取器。</summary>
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

    /// <summary>使用 CsvHelper 的异步解析器读取 CSV 记录。</summary>
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

/// <summary>对不可定位的 CSV 源流施加读取字节上限且不拥有底层流的包装器。</summary>
internal sealed class CsvLimitedReadStream : Stream
{
    private readonly Stream _inner;
    private readonly long _maxBytes;
    private long _readBytes;

    /// <summary>创建对底层输入流实施字节上限的包装器。</summary>
    public CsvLimitedReadStream(Stream inner, long maxBytes)
    {
        _inner = inner;
        _maxBytes = maxBytes;
    }

    public override bool CanRead => _inner.CanRead;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

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

    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { }
}

/// <summary>
/// 将调用方的取消令牌绑定到 CsvHelper/StreamReader 未提供令牌参数的异步入口。
/// </summary>
internal sealed class CsvCancellationStream : Stream
{
    private readonly Stream _inner;
    private readonly CancellationToken _cancellationToken;
    private readonly bool _suppressSynchronousFlush;

    public CsvCancellationStream(Stream inner, CancellationToken cancellationToken,
        bool suppressSynchronousFlush = false)
    {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _cancellationToken = cancellationToken;
        _suppressSynchronousFlush = suppressSynchronousFlush;
    }

    public override bool CanRead => _inner.CanRead;
    public override bool CanSeek => _inner.CanSeek;
    public override bool CanWrite => _inner.CanWrite;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => _inner.Position = value; }

    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        var read = await _inner.ReadAsync(buffer, offset, count, _cancellationToken).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
        return read;
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        _inner.Write(buffer, offset, count);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    public override async Task WriteAsync(byte[] buffer, int offset, int count,
        CancellationToken cancellationToken)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        await _inner.WriteAsync(buffer, offset, count, _cancellationToken).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    public override void Flush()
    {
        if (!_suppressSynchronousFlush)
            _inner.Flush();
    }

    public override async Task FlushAsync(CancellationToken cancellationToken)
    {
        _cancellationToken.ThrowIfCancellationRequested();
        await _inner.FlushAsync(_cancellationToken).ConfigureAwait(false);
        _cancellationToken.ThrowIfCancellationRequested();
    }

    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    public override void SetLength(long value) => _inner.SetLength(value);

    protected override void Dispose(bool disposing)
    {
        // The caller owns the underlying stream.
        base.Dispose(disposing);
    }
}

/// <summary>基于 CsvHelper 的 CSV 记录写入器。</summary>
internal sealed class CsvRecordWriter : IDisposable
{
    private readonly CsvWriter _csv;
    private readonly CsvFormulaInjectionPolicy _formulaInjectionPolicy;

    /// <summary>为一次导出创建并复用 CsvHelper 写入器。</summary>
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

    /// <summary>写入一个字段，并按请求策略保护潜在公式。</summary>
    public void WriteField(string field, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _csv.WriteField(ProtectFormula(field ?? string.Empty, _formulaInjectionPolicy));
    }

    /// <summary>结束当前记录。</summary>
    public void NextRecord() => _csv.NextRecord();

    /// <summary>异步结束当前记录。</summary>
    public async Task NextRecordAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.NextRecordAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>刷新 CsvHelper 和底层文本写入器。</summary>
    public void Flush() => _csv.Flush();

    /// <summary>异步刷新 CsvHelper 和底层文本写入器。</summary>
    public async Task FlushAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.FlushAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <inheritdoc />
    public void Dispose() => _csv.Dispose();

    internal static string ProtectFormula(string value, CsvFormulaInjectionPolicy policy)
    {
        if (policy != CsvFormulaInjectionPolicy.Escape || !StartsWithFormula(value))
            return value;
        return $"'{value}";
    }

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

/// <summary>在内存中完成 CSV 格式化，并通过异步字节写入提交记录。</summary>
internal sealed class CsvAsyncRecordWriter : IDisposable
{
    private readonly Stream _destination;
    private readonly Encoding _encoding;
    private readonly Encoder _encoder;
    private readonly StringBuilder _buffer;
    private readonly StringWriter _textWriter;
    private readonly CsvWriter _csv;
    private readonly CsvFormulaInjectionPolicy _formulaInjectionPolicy;
    private readonly byte[] _byteBuffer = new byte[4096];
    private bool _preambleWritten;
    private bool _disposed;

    /// <summary>初始化异步 CSV 记录写入器。</summary>
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

    /// <summary>写入一个字段；该操作只修改内存中的 CSV 文本缓冲。</summary>
    public void WriteField(string field, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _csv.WriteField(CsvRecordWriter.ProtectFormula(field ?? string.Empty, _formulaInjectionPolicy));
    }

    /// <summary>完成当前记录，并通过目标流的异步写入接口提交。</summary>
    public async Task NextRecordAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await _csv.NextRecordAsync().ConfigureAwait(false);
        await WriteBufferedTextAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>刷新 CSV 编码器和目标流，整个 IO 路径使用异步接口。</summary>
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
