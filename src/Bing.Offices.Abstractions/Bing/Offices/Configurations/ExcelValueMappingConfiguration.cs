namespace Bing.Offices.Configurations;

/// <summary>
/// 可序列化的 Excel 显示值映射。
/// </summary>
public sealed class ExcelValueMappingConfiguration
{
    /// <summary>
    /// 获取或设置单元格显示文本。
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// 获取或设置实体属性文本值。
    /// </summary>
    public string Value { get; set; }
}
