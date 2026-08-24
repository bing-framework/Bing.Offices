namespace Bing.Offices.Validations;

/// <summary>
/// 映射计划在构建阶段绑定的只读校验描述。
/// </summary>
public interface IExcelValidationBinding
{
    /// <summary>获取规则类型。</summary>
    ExcelValidationBindingKind Kind { get; }

    /// <summary>获取是否在类型转换前执行。</summary>
    bool IsRaw { get; }

    /// <summary>获取失败消息。</summary>
    string ErrorMessage { get; }

    /// <summary>执行当前绑定。</summary>
    bool Validate(ExcelValidationContext context);
}

/// <summary>
/// 预绑定校验规则类型。
/// </summary>
public enum ExcelValidationBindingKind
{
    /// <summary>普通或未知规则。</summary>
    Custom,
    /// <summary>必填规则。</summary>
    Required,
    /// <summary>正则规则。</summary>
    Regex,
    /// <summary>日期规则。</summary>
    Date,
    /// <summary>最大值规则。</summary>
    MaxValue,
    /// <summary>范围规则。</summary>
    Range,
    /// <summary>最大长度规则。</summary>
    MaxLength,
    /// <summary>唯一性规则。</summary>
    Unique
}
