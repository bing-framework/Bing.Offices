namespace Bing.Offices.Conversions;

/// <summary>
/// 单元格值转换器
/// </summary>
public interface ICellValueConverter
{
    /// <summary>
    /// 获取单元格字符串值
    /// </summary>
    /// <param name="cell">单元格</param>
    string GetStringValue(object cell);
}