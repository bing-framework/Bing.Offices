using System.Globalization;
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

    public override void Flush() => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    protected override void Dispose(bool disposing) { }
}

/// <summary>基于 CsvHelper 的 CSV 记录写入器。</summary>
internal static class CsvRecordWriter
{
    /// <summary>将一个记录写入调用方拥有的文本写入器。</summary>
    public static void Write(TextWriter writer, IEnumerable<string> fields, char delimiter, char quote, string newLine,
        CsvFormulaInjectionPolicy formulaInjectionPolicy)
    {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            NewLine = newLine
        };
        using var csv = new CsvWriter(writer, configuration, true);
        foreach (var field in fields)
            csv.WriteField(ProtectFormula(field ?? string.Empty, formulaInjectionPolicy));
        csv.NextRecord();
    }

    private static string ProtectFormula(string value, CsvFormulaInjectionPolicy policy)
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
