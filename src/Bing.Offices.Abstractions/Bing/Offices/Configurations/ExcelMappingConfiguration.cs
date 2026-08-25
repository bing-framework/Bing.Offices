namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 请求级映射配置。
/// </summary>
public sealed class ExcelMappingConfiguration
{
    /// <summary>
    /// 获取或设置该方向使用的业务 Profile 名称。
    /// </summary>
    public string Profile { get; set; }

    /// <summary>
    /// 获取或设置该方向使用的稳定模型别名。
    /// </summary>
    public string ModelAlias { get; set; }

    /// <summary>
    /// 获取或设置是否清空低优先级动态列。
    /// </summary>
    public bool ClearDynamicColumns { get; set; }

    /// <summary>
    /// 获取或设置需要从低优先级动态列中移除的稳定 Key 集合。
    /// </summary>
    public List<string> DynamicColumnKeysToRemove { get; set; } = new List<string>();

    /// <summary>
    /// 获取或设置动态列集合的合并方式。
    /// </summary>
    public ExcelDynamicColumnMergeMode? DynamicColumnMergeMode { get; set; }

    /// <summary>
    /// 获取或设置是否重置低优先级样式配置。
    /// </summary>
    public bool ResetStyle { get; set; }

    /// <summary>
    /// 获取或设置是否重置低优先级布局配置。
    /// </summary>
    public bool ResetLayout { get; set; }

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
