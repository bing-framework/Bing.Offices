using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

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
