using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;

namespace Bing.Offices.Validations;

/// <summary>
/// 内置 Excel 导入校验规则集合。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class ExcelValidationRules
{
    /// <summary>
    /// 创建无状态内置校验规则。
    /// </summary>
    /// <returns>按固定执行顺序排列的内置校验规则集合。</returns>
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
