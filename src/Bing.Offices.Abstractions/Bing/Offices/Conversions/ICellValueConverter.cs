namespace Bing.Offices.Conversions;

/// <summary>
/// 旧版单元格文本转换器兼容契约。
/// </summary>
/// <remarks>
/// 该契约仅用于迁移已有的文本提取逻辑。新代码应实现 <see cref="IExcelValueConverter"/>，
/// 不得依赖具体 Workbook 生命周期。
/// </remarks>
[Obsolete("请实现 IExcelValueConverter 以获得双向、提供程序无关的转换能力。")]
public interface ICellValueConverter
{
    /// <summary>
    /// 从提供程序单元格对象提取文本。
    /// </summary>
    /// <param name="cell">提供程序单元格对象。</param>
    string GetStringValue(object cell);
}
