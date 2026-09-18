using Bing.Offices.Attributes;
using Bing.Offices.Validations;

namespace Bing.Offices.Mappings;

/// <summary>
/// 将特性或命名规则绑定为统一校验执行单元。
/// </summary>
internal sealed class ExcelValidationBinding : IExcelValidationBinding
{
    /// <summary>
    /// 用于特性校验绑定的校验属性实例。
    /// </summary>
    private readonly FilterAttributeBase _attribute;
    /// <summary>
    /// 执行特性校验的规则实例。
    /// </summary>
    private readonly IExcelValidationRule _attributeRule;
    /// <summary>
    /// 执行命名校验的规则实例。
    /// </summary>
    private readonly INamedExcelValidationRule _namedRule;

    /// <summary>
    /// 初始化一个 <see cref="ExcelValidationBinding" /> 类型的实例。
    /// </summary>
    /// <param name="attribute">声明校验配置的属性特性。</param>
    /// <param name="attributeRule">执行该特性校验的规则。</param>
    private ExcelValidationBinding(FilterAttributeBase attribute, IExcelValidationRule attributeRule)
    {
        _attribute = attribute;
        _attributeRule = attributeRule;
        Kind = ResolveKind(attribute);
        IsRaw = attribute is ExcelRequiredAttribute || attribute is ExcelRegexAttribute;
        ErrorMessage = attribute.ErrorMsg;
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelValidationBinding" /> 类型的实例。
    /// </summary>
    /// <param name="namedRule">按配置名称解析得到的命名校验规则。</param>
    private ExcelValidationBinding(INamedExcelValidationRule namedRule)
    {
        _namedRule = namedRule ?? throw new ArgumentNullException(nameof(namedRule));
        Kind = ExcelValidationBindingKind.Custom;
        ErrorMessage = namedRule.ErrorMessage;
    }

    /// <summary>
    /// 创建由属性特性和对应规则组成的校验绑定。
    /// </summary>
    /// <param name="attribute">声明校验配置的属性特性。</param>
    /// <param name="rule">执行该特性校验的规则。</param>
    /// <returns>绑定特性与规则后的校验执行单元。</returns>
    internal static ExcelValidationBinding Attribute(FilterAttributeBase attribute, IExcelValidationRule rule)
    {
        if (attribute == null)
            throw new ArgumentNullException(nameof(attribute));
        if (rule == null)
            throw new ArgumentNullException(nameof(rule));
        return new ExcelValidationBinding(attribute, rule);
    }

    /// <summary>
    /// 创建由命名规则组成的校验绑定。
    /// </summary>
    /// <param name="rule">按配置名称解析得到的命名校验规则。</param>
    /// <returns>绑定命名规则后的校验执行单元。</returns>
    internal static ExcelValidationBinding Named(INamedExcelValidationRule rule) =>
        new ExcelValidationBinding(rule);

    /// <inheritdoc />
    public ExcelValidationBindingKind Kind { get; }
    /// <inheritdoc />
    public bool IsRaw { get; }
    /// <inheritdoc />
    public string ErrorMessage { get; }

    /// <inheritdoc />
    public bool Validate(ExcelValidationContext context) => _attribute != null
        ? _attributeRule.Validate(_attribute, context)
        : _namedRule.Validate(context);

    /// <summary>
    /// 将校验特性映射为统一的绑定类型。
    /// </summary>
    /// <param name="attribute">待分类的校验特性。</param>
    /// <returns>对应的校验绑定类型。</returns>
    private static ExcelValidationBindingKind ResolveKind(FilterAttributeBase attribute) => attribute switch
    {
        ExcelRequiredAttribute => ExcelValidationBindingKind.Required,
        ExcelRegexAttribute => ExcelValidationBindingKind.Regex,
        ExcelDateAttribute => ExcelValidationBindingKind.Date,
        ExcelMaxValueAttribute => ExcelValidationBindingKind.MaxValue,
        ExcelRangeAttribute => ExcelValidationBindingKind.Range,
        ExcelMaxLengthAttribute => ExcelValidationBindingKind.MaxLength,
        ExcelUniqueAttribute => ExcelValidationBindingKind.Unique,
        _ => ExcelValidationBindingKind.Custom
    };
}
