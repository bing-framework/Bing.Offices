using Bing.Offices.Attributes;

namespace Bing.Offices.Validations;

/// <summary>
/// Excel 导入属性校验规则。
/// </summary>
public interface IExcelValidationRule
{
    /// <summary>
    /// 判断规则是否支持指定特性。
    /// </summary>
    /// <param name="attribute">校验特性。</param>
    /// <returns>支持该特性时为 <see langword="true"/>，否则为 <see langword="false"/>。</returns>
    bool CanValidate(FilterAttributeBase attribute);

    /// <summary>
    /// 校验当前单元格值。
    /// </summary>
    /// <param name="attribute">校验特性。</param>
    /// <param name="context">校验上下文。</param>
    /// <returns>校验通过时为 <see langword="true"/>，否则为 <see langword="false"/>。</returns>
    bool Validate(FilterAttributeBase attribute, ExcelValidationContext context);
}
