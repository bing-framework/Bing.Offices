using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;

namespace Bing.Offices.Validations;

/// <summary>
/// 重复值校验规则。
/// </summary>
internal sealed class DuplicationExcelValidationRule : IExcelValidationRule
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
