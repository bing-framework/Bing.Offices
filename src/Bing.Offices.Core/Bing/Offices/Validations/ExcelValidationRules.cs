using System.Collections.Concurrent;
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
    public bool CanValidate(FilterAttributeBase attribute) => attribute is RequiredAttribute
        || attribute is ExcelRequiredAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
        !string.IsNullOrWhiteSpace(context.Value);
}

/// <summary>
/// 正则表达式校验规则。
/// </summary>
public sealed class RegexExcelValidationRule : IExcelValidationRule
{
    private static readonly ConcurrentDictionary<string, Regex> RegexCache = new(StringComparer.Ordinal);

    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is RegexAttribute
        || attribute is ExcelRegexAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var pattern = attribute is RegexAttribute legacy ? legacy.RegexString : ((ExcelRegexAttribute)attribute).Pattern;
        var regex = RegexCache.GetOrAdd(pattern, value => new Regex(value, RegexOptions.Compiled,
            TimeSpan.FromSeconds(1)));
        return regex.IsMatch(context.Value);
    }
}

/// <summary>
/// 数值区间校验规则。
/// </summary>
public sealed class RangeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is RangeAttribute
        || attribute is ExcelRangeAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var minimum = attribute is RangeAttribute legacy ? legacy.Min : ((ExcelRangeAttribute)attribute).Min;
        var maximum = attribute is RangeAttribute old ? old.Max : ((ExcelRangeAttribute)attribute).Max;
        var valueText = Convert.ToString(context.ConvertedValue ?? context.Value, context.Culture);
        return decimal.TryParse(valueText, NumberStyles.Number, context.Culture, out var value)
               && value >= Convert.ToDecimal(minimum, CultureInfo.InvariantCulture)
               && value <= Convert.ToDecimal(maximum, CultureInfo.InvariantCulture);
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
    public bool CanValidate(FilterAttributeBase attribute) => attribute is MaxLengthAttribute
        || attribute is ExcelMaxLengthAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var maxLength = attribute is MaxLengthAttribute legacy
            ? legacy.MaxLength
            : ((ExcelMaxLengthAttribute)attribute).MaxLength;
        return context.Value.Length <= maxLength;
    }
}

/// <summary>
/// 日期校验规则。
/// </summary>
public sealed class DateTimeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is DateTimeAttribute
        || attribute is ExcelDateAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
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
        var v2 = attribute as ExcelDateAttribute;
        var culture = context.Culture;
        if (!string.IsNullOrWhiteSpace(v2?.CultureName))
            culture = CultureInfo.GetCultureInfo(v2.CultureName);
        return !string.IsNullOrWhiteSpace(v2?.Format)
            ? DateTime.TryParseExact(context.Value, v2.Format, culture, DateTimeStyles.None, out _)
            : DateTime.TryParse(context.Value, culture, DateTimeStyles.AllowWhiteSpaces, out _);
    }
}

/// <summary>
/// 重复值校验规则。
/// </summary>
public sealed class DuplicationExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is DuplicationAttribute
        || attribute is ExcelUniqueAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        // Unique 需要跨单元格维护 committed/pending 状态，由执行器的 UniqueTracker 统一处理。
        return true;
    }
}
