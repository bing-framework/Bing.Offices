namespace Bing.Offices.Styles;

/// <summary>
/// 单个边框的提供程序无关描述。
/// </summary>
public sealed class ExcelBorderStyle
{
    /// <summary>
    /// 获取或初始化线型。
    /// </summary>
    public ExcelBorderLineStyle LineStyle { get; init; }

    /// <summary>
    /// 获取或初始化线条颜色。
    /// </summary>
    public ExcelColor Color { get; init; }
}
