using Bing.Offices.Styles;

// ReSharper disable once CheckNamespace
namespace Bing.Offices.Attributes;

/// <summary>
/// 配置实体类型导出时的表头字体样式。
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class HeaderAttribute : DecoratorAttributeBase
{
    /// <summary>
    /// 获取或设置表头字体颜色。
    /// </summary>
    public Color Color { get; set; } = Color.Black;

    /// <summary>
    /// 获取或设置表头字体名称。
    /// </summary>
    public string FontName { get; set; } = "微软雅黑";

    /// <summary>
    /// 获取或设置表头字体大小。
    /// </summary>
    public int FontSize { get; set; } = 12;

    /// <summary>
    /// 获取或设置表头字体是否加粗；为 <see langword="true" /> 时使用粗体。
    /// </summary>
    public bool Bold { get; set; } = true;
}
