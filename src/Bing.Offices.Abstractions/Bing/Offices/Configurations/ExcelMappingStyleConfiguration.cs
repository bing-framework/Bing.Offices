namespace Bing.Offices.Configurations;

/// <summary>
/// 规范化映射文档中的跨提供程序样式描述。
/// </summary>
public sealed class ExcelMappingStyleConfiguration
{
    /// <summary>获取或设置表头样式键。</summary>
    public string HeaderStyleKey { get; set; }

    /// <summary>获取或设置是否清除低优先级表头样式键。</summary>
    public bool ClearHeaderStyleKey { get; set; }

    /// <summary>获取或设置正文样式键。</summary>
    public string BodyStyleKey { get; set; }

    /// <summary>获取或设置是否清除低优先级正文样式键。</summary>
    public bool ClearBodyStyleKey { get; set; }

    /// <summary>获取或设置数字格式。</summary>
    public string NumberFormat { get; set; }

    /// <summary>获取或设置是否清除低优先级数字格式。</summary>
    public bool ClearNumberFormat { get; set; }
}
