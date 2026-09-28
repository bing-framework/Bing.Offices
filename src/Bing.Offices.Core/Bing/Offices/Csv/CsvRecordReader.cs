using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

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
