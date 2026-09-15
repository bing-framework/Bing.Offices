namespace Bing.Offices.Styles;

/// <summary>
/// 与 Excel 提供程序无关的颜色描述。
/// </summary>
public sealed class ExcelColor
{
    /// <summary>
    /// 获取或初始化ARGB 十六进制颜色，例如 <c>FF1F4E79</c>。
    /// </summary>
    public string Argb { get; init; }

    /// <summary>初始化一个 <see cref="ExcelColor" /> 类型的实例。</summary>
    public ExcelColor()
    {
    }

    /// <summary>初始化一个 <see cref="ExcelColor" /> 类型的实例。</summary>
    /// <param name="argb">ARGB 十六进制颜色文本。</param>
    public ExcelColor(string argb) => Argb = argb;
}

