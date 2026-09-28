using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;

namespace Bing.Offices.Validations;

/// <summary>
/// 正则表达式校验规则。
/// </summary>
internal sealed class RegexExcelValidationRule : IExcelValidationRule
{
    /// <summary>
    /// 进程级正则缓存允许保留的最大模式数量（256 个）。
    /// </summary>
    internal const int RegexCacheCapacity = 256;
    /// <summary>
    /// 保护正则缓存和淘汰队列的一致性锁。
    /// </summary>
    private static readonly object RegexCacheLock = new object();
    /// <summary>
    /// 按模式文本缓存的已编译正则表达式。
    /// </summary>
    private static readonly Dictionary<string, Regex> RegexCache = new Dictionary<string, Regex>(StringComparer.Ordinal);
    /// <summary>
    /// 按插入顺序记录缓存模式，用于有界先进先出淘汰。
    /// </summary>
    private static readonly Queue<string> RegexCacheOrder = new Queue<string>();

    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelRegexAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var pattern = ((ExcelRegexAttribute)attribute).Pattern;
        var regex = GetRegex(pattern);
        return regex.IsMatch(context.Value);
    }

    /// <summary>
    /// 获取带容量上限的正则表达式实例，避免用户输入大量不同模式导致进程级缓存无界增长。
    /// </summary>
    /// <param name="pattern">正则表达式模式。</param>
    /// <returns>可复用的正则表达式实例。</returns>
    private static Regex GetRegex(string pattern)
    {
        lock (RegexCacheLock)
        {
            if (RegexCache.TryGetValue(pattern, out var regex))
                return regex;
            regex = new Regex(pattern, RegexOptions.Compiled, TimeSpan.FromSeconds(1));
            RegexCache[pattern] = regex;
            RegexCacheOrder.Enqueue(pattern);
            while (RegexCache.Count > RegexCacheCapacity)
                RegexCache.Remove(RegexCacheOrder.Dequeue());
            return regex;
        }
    }
}
