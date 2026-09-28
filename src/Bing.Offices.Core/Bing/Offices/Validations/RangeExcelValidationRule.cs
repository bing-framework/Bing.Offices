using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;

namespace Bing.Offices.Validations;

/// <summary>
/// 数值区间校验规则。
/// </summary>
internal sealed class RangeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelRangeAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var range = (ExcelRangeAttribute)attribute;
        var valueText = Convert.ToString(context.ConvertedValue ?? context.Value, context.Culture);
        return decimal.TryParse(valueText, NumberStyles.Number, context.Culture, out var value)
             && value >= Convert.ToDecimal(range.Min, CultureInfo.InvariantCulture)
             && value <= Convert.ToDecimal(range.Max, CultureInfo.InvariantCulture);
    }
}
