namespace Bing.Offices.Configurations;

/// <summary>
/// 规范化映射文档中的跨提供程序布局描述。
/// </summary>
public sealed class ExcelMappingLayoutConfiguration
{
    /// <summary>获取或设置物理列索引。</summary>
    public int? ColumnIndex { get; set; }

    /// <summary>获取或设置是否重置低优先级物理列索引。</summary>
    public bool ResetColumnIndex { get; set; }

    /// <summary>获取或设置相对布局键。</summary>
    public string PlacementKey { get; set; }

    /// <summary>获取或设置是否清除低优先级相对布局键。</summary>
    public bool ClearPlacementKey { get; set; }
}
