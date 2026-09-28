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
