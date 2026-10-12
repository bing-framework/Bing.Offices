using System;
using Bing.Offices.Conversions;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Business;

/// <summary>
/// 将业务渠道代码转换为透视来源展示值。
/// </summary>
internal sealed class NpoiPivotBusinessChannelConverter : INamedExcelValueConverter
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotBusinessChannelConverter" /> 类型的实例。
    /// </summary>
    public NpoiPivotBusinessChannelConverter() { }

    /// <inheritdoc />
    public string Name => "pivot-business-channel";

    /// <inheritdoc />
    public bool CanConvert(Type propertyType) => propertyType == typeof(string);

    /// <inheritdoc />
    public bool TryConvertFrom(ExcelConversionContext context, out object value)
    {
        value = context.Value;
        return false;
    }

    /// <inheritdoc />
    public bool TryConvertTo(ExcelConversionContext context, out object value)
    {
        value = context.Value switch
        {
            "R" => "Retail",
            "W" => "Wholesale",
            _ => context.Value
        };
        return true;
    }
}
