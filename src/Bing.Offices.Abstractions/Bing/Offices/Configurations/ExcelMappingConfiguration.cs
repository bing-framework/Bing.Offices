namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 请求级映射配置。
/// </summary>
public sealed class ExcelMappingConfiguration
{
    /// <summary>
    /// 获取或设置列配置集合。
    /// </summary>
    public List<ExcelColumnConfiguration> Columns { get; set; } = new List<ExcelColumnConfiguration>();
}
