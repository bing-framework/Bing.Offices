using System;
using Bing.Offices.Validations;

namespace Bing.Offices.Mappings;

/// <summary>
/// 统一创建预绑定的命名校验描述。
/// </summary>
public static class ExcelValidationBindingFactory
{
    /// <summary>
    /// 创建命名校验绑定。
    /// </summary>
    public static IExcelValidationBinding CreateNamed(INamedExcelValidationRule rule)
    {
        if (rule == null)
            throw new ArgumentNullException(nameof(rule));
        return ExcelValidationBinding.Named(rule);
    }
}
