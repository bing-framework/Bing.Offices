namespace Bing.Offices.Configurations;

/// <summary>
/// normalized mapping document 中的 provider-neutral 布局描述。
/// </summary>
public sealed class ExcelMappingLayoutConfiguration
{
    /// <summary>获取或设置物理列索引。</summary>
    public int? ColumnIndex { get; set; }
    /// <summary>获取或设置相对布局键。</summary>
    public string PlacementKey { get; set; }
}
