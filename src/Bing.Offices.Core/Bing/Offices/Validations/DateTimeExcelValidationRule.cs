using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;

namespace Bing.Offices.Validations;

/// <summary>
/// 日期校验规则。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class DateTimeExcelValidationRule : IExcelValidationRule
{
    /// <inheritdoc />
    public bool CanValidate(FilterAttributeBase attribute) => attribute is ExcelDateAttribute;

    /// <inheritdoc />
    public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context)
    {
        var attributeValue = (ExcelDateAttribute)attribute;
        return TryParseValue(context.Cell, context.Value, context.PropertyType ?? typeof(DateTime),
            context.Culture, attributeValue, out _);
    }

    /// <summary>
    /// 尝试按统一日期转换合同解析日期值，供 Provider 转换边界复用。
    /// </summary>
    /// <param name="cell">原始单元格值。</param>
    /// <param name="text">规范化文本。</param>
    /// <param name="targetType">目标日期类型。</param>
    /// <param name="culture">请求区域性。</param>
    /// <param name="attribute">日期输入配置。</param>
    /// <param name="value">解析后的日期值。</param>
    /// <returns>解析成功时为 true，并通过 <paramref name="value" /> 返回日期值；失败时为 false。</returns>
    public bool TryParseValue(ExcelCellValue cell, string text, Type targetType, CultureInfo culture,
        ExcelDateAttribute attribute, out object value) =>
        ExcelDateParser.TryParse(cell, text, targetType, culture, attribute, out value);

    /// <summary>
    /// 按工作簿数据校验语义解析日期或时间值，供 Provider 校验边界复用。
    /// </summary>
    /// <param name="cell">原始单元格值。</param>
    /// <param name="text">单元格或约束的文本值。</param>
    /// <param name="timeOnly">是否只解析时间部分。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="value">解析后的无时区日期时间。</param>
    /// <returns>解析成功时为 true，失败时为 false。</returns>
    public bool TryParseWorkbookDate(ExcelCellValue cell, string text, bool timeOnly, bool isDate1904,
        out DateTime value) =>
        ExcelDateParser.TryParseValidation(cell, text, timeOnly, isDate1904, out value);
}
