namespace Bing.Offices.Conversions;

/// <summary>
/// 可通过配置名称选择的 Excel 值转换器。
/// </summary>
public interface INamedExcelValueConverter : IExcelValueConverter
{
    /// <summary>
    /// 获取配置中使用的唯一名称。
    /// </summary>
    string Name { get; }
}
