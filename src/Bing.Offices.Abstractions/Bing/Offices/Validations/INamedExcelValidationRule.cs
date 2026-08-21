namespace Bing.Offices.Validations;

/// <summary>
/// 可通过配置名称选择的 Excel 校验规则。
/// </summary>
public interface INamedExcelValidationRule
{
    /// <summary>
    /// 获取配置中使用的唯一名称。
    /// </summary>
    string Name { get; }

    /// <summary>
    /// 获取校验失败时的错误信息。
    /// </summary>
    string ErrorMessage { get; }

    /// <summary>
    /// 校验当前单元格值。
    /// </summary>
    /// <param name="context">校验上下文。</param>
    bool Validate(ExcelValidationContext context);
}
