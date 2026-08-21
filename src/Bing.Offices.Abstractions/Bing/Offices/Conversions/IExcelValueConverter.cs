namespace Bing.Offices.Conversions;

/// <summary>
/// Excel 属性值双向转换器。
/// </summary>
public interface IExcelValueConverter
{
    /// <summary>
    /// 判断转换器是否支持指定属性类型。
    /// </summary>
    /// <param name="propertyType">属性类型。</param>
    bool CanConvert(Type propertyType);

    /// <summary>
    /// 尝试将导入单元格值转换为属性值。
    /// </summary>
    /// <param name="context">转换上下文。</param>
    /// <param name="value">转换后的属性值。</param>
    /// <returns>已处理时返回 <see langword="true"/>；否则使用默认转换。</returns>
    bool TryConvertFrom(ExcelConversionContext context, out object value);

    /// <summary>
    /// 尝试将属性值转换为导出单元格值。
    /// </summary>
    /// <param name="context">转换上下文。</param>
    /// <param name="value">转换后的单元格值。</param>
    /// <returns>已处理时返回 <see langword="true"/>；否则使用默认转换。</returns>
    bool TryConvertTo(ExcelConversionContext context, out object value);
}
