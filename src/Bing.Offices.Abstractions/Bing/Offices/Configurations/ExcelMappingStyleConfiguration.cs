namespace Bing.Offices.Configurations;

/// <summary>
/// normalized mapping document 中的 provider-neutral 样式描述。
/// </summary>
public sealed class ExcelMappingStyleConfiguration
{
    /// <summary>获取或设置表头样式键。</summary>
    public string HeaderStyleKey { get; set; }
    /// <summary>获取或设置正文样式键。</summary>
    public string BodyStyleKey { get; set; }
    /// <summary>获取或设置数字格式。</summary>
    public string NumberFormat { get; set; }
}
