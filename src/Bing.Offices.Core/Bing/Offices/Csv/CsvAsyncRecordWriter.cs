using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

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
