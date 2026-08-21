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
    public bool CanValidate(FilterAttributeBase attribute) => attribute is RequiredAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
        !string.IsNullOrWhiteSpace(context.Value);
}

/// <summary>
/// 正则表达式校验规则。
/// </summary>
public sealed class RegexExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is RegexAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
        Regex.IsMatch(context.Value, ((RegexAttribute)attribute).RegexString);
}

/// <summary>
/// 数值区间校验规则。
/// </summary>
public sealed class RangeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is RangeAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var range = (RangeAttribute)attribute;
        var valueText = Convert.ToString(context.ConvertedValue ?? context.Value, context.Culture);
        return decimal.TryParse(valueText, NumberStyles.Number, context.Culture, out var value)
               && value >= range.Min && value <= range.Max;
    }
}

/// <summary>
/// 最大长度校验规则。
/// </summary>
public sealed class MaxLengthExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is MaxLengthAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
        context.Value.Length <= ((MaxLengthAttribute)attribute).MaxLength;
}

/// <summary>
/// 日期校验规则。
/// </summary>
public sealed class DateTimeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is DateTimeAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
        context.ConvertedValue is DateTime || DateTime.TryParse(context.Value, context.Culture,
            DateTimeStyles.None, out _);
}

/// <summary>
/// 重复值校验规则。
/// </summary>
public sealed class DuplicationExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is DuplicationAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        if (string.IsNullOrWhiteSpace(context.Value))
            return true;
        if (!context.DuplicateValues.TryGetValue(context.PropertyName, out var values))
        {
            values = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            context.DuplicateValues[context.PropertyName] = values;
        }
        return values.Add(context.Value);
    }
}
