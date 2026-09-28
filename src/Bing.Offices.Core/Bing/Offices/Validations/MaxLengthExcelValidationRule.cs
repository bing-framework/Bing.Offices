using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;

namespace Bing.Offices.Validations;

/// <summary>
/// 最大长度校验规则。
/// </summary>
internal sealed class MaxLengthExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelMaxLengthAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var maxLength = ((ExcelMaxLengthAttribute)attribute).MaxLength;
        return context.Value.Length <= maxLength;
    }
}
