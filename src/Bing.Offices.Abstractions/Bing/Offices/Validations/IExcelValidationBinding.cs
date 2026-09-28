namespace Bing.Offices.Validations;

/// <summary>
/// 映射计划在构建阶段绑定的只读校验描述。
/// </summary>
public interface IExcelValidationBinding
{
    /// <summary>
    /// 获取规则类型。
    /// </summary>
    ExcelValidationBindingKind Kind { get; }

    /// <summary>
    /// 获取是否在类型转换前执行。
    /// </summary>
    bool IsRaw { get; }

    /// <summary>
    /// 获取失败消息。
    /// </summary>
    string ErrorMessage { get; }

    /// <summary>
    /// 执行当前绑定。
    /// </summary>
    /// <param name="context">当前单元格的校验上下文。</param>
    /// <returns>校验通过时为 <see langword="true"/>，否则为 <see langword="false"/>。</returns>
    bool Validate(ExcelValidationContext context);
}
