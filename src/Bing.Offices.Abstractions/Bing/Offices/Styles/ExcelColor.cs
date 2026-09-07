namespace Bing.Offices.Styles;

/// <summary>
/// 与 Excel 提供程序无关的颜色描述。
/// </summary>
public sealed class ExcelColor
{
    /// <summary>
    /// 获取或设置 ARGB 十六进制颜色，例如 <c>FF1F4E79</c>。
    /// </summary>
    public string Argb { get; init; }

    /// <summary>
    /// 创建颜色描述。
    /// </summary>
    public ExcelColor()
    {
    }

    /// <summary>
    /// 创建颜色描述。
    /// </summary>
    public ExcelColor(string argb) => Argb = argb;
}

