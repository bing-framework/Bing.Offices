using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

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
