using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;

namespace Bing.Offices.Validations;

/// <summary>
/// 内置 Excel 导入校验规则集合。
/// </summary>
public static class ExcelValidationRules
{
    /// <summary>
    /// 创建无状态内置校验规则。
    /// </summary>
    public static IReadOnlyList<IExcelValidationRule> CreateDefault() => new IExcelValidationRule[]
    {
        new RequiredExcelValidationRule(),
        new RegexExcelValidationRule(),
        new RangeExcelValidationRule(),
        new MaxValueExcelValidationRule(),
        new MaxLengthExcelValidationRule(),
        new DateTimeExcelValidationRule(),
        new DuplicationExcelValidationRule()
    };
}

/// <summary>
/// 必填校验规则。
/// </summary>
public sealed class RequiredExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelRequiredAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
        !string.IsNullOrWhiteSpace(context.Value);
}

/// <summary>
/// 正则表达式校验规则。
/// </summary>
public sealed class RegexExcelValidationRule : IExcelValidationRule
{
    /// <summary>进程级正则缓存允许保留的最大模式数量。</summary>
    internal const int RegexCacheCapacity = 256;
    /// <summary>保护正则缓存和淘汰队列的一致性锁。</summary>
    private static readonly object RegexCacheLock = new object();
    /// <summary>按模式文本缓存的已编译正则表达式。</summary>
    private static readonly Dictionary<string, Regex> RegexCache = new Dictionary<string, Regex>(StringComparer.Ordinal);
    /// <summary>按插入顺序记录缓存模式，用于有界先进先出淘汰。</summary>
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

/// <summary>
/// 数值区间校验规则。
/// </summary>
public sealed class RangeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelRangeAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
         var range = (ExcelRangeAttribute)attribute;
        var valueText = Convert.ToString(context.ConvertedValue ?? context.Value, context.Culture);
        return decimal.TryParse(valueText, NumberStyles.Number, context.Culture, out var value)
             && value >= Convert.ToDecimal(range.Min, CultureInfo.InvariantCulture)
             && value <= Convert.ToDecimal(range.Max, CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// 最大值校验规则。
/// </summary>
public sealed class MaxValueExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelMaxValueAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var valueText = Convert.ToString(context.ConvertedValue ?? context.Value, context.Culture);
        return decimal.TryParse(valueText, NumberStyles.Number, context.Culture, out var value)
               && value <= Convert.ToDecimal(((ExcelMaxValueAttribute)attribute).MaxValue,
                   CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// 最大长度校验规则。
/// </summary>
public sealed class MaxLengthExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelMaxLengthAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var maxLength = ((ExcelMaxLengthAttribute)attribute).MaxLength;
        return context.Value.Length <= maxLength;
    }
}

/// <summary>
/// 日期校验规则。
/// </summary>
public sealed class DateTimeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelDateAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var v2 = attribute as ExcelDateAttribute;
        var culture = context.Culture;
        if (!string.IsNullOrWhiteSpace(v2?.CultureName))
            culture = CultureInfo.GetCultureInfo(v2.CultureName);
        if (!string.IsNullOrWhiteSpace(v2?.Format))
            return DateTime.TryParseExact(context.Value, v2.Format, culture, DateTimeStyles.None, out _);
        if (context.ConvertedValue is DateTime || context.ConvertedValue is DateTimeOffset)
            return true;
        if (context.Cell?.Value is DateTime || context.Cell?.Value is DateTimeOffset)
            return true;
        if (context.Cell?.Value is double serial)
        {
            try
            {
                DateTime.FromOADate(serial);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
        return DateTime.TryParse(context.Value, culture, DateTimeStyles.AllowWhiteSpaces, out _);
    }
}

/// <summary>
/// 重复值校验规则。
/// </summary>
public sealed class DuplicationExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelUniqueAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        // Unique 需要跨单元格维护 committed/pending 状态，由执行器的 UniqueTracker 统一处理。
        return true;
    }
}
