namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 请求级映射配置。
/// </summary>
public sealed class ExcelMappingConfiguration
{
    /// <summary>
    /// 获取或设置该配置的来源。
    /// </summary>
    public MappingSourceKind SourceKind { get; set; } = MappingSourceKind.Request;

    /// <summary>
    /// 获取或设置列配置集合。
    /// </summary>
    public List<ExcelColumnConfiguration> Columns { get; set; } = new List<ExcelColumnConfiguration>();

    /// <summary>获取或设置规范化动态列描述。</summary>
    public List<ExcelMappingDynamicColumnConfiguration> DynamicColumns { get; set; } =
        new List<ExcelMappingDynamicColumnConfiguration>();

    /// <summary>获取或设置规范化样式描述。</summary>
    public ExcelMappingStyleConfiguration Style { get; set; }

    /// <summary>获取或设置规范化布局描述。</summary>
    public ExcelMappingLayoutConfiguration Layout { get; set; }
}
