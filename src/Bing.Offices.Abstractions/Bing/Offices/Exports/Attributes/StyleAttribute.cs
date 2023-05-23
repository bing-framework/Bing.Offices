using Bing.Offices.Enums.Styles;

namespace Bing.Offices.Exports.Attributes;

/// <summary>
/// 样式 特性
/// </summary>
public class StyleAttribute: Attribute
{
    /// <summary>
    /// 样式
    /// </summary>
    public IBaseStyle Style { get; set; }

    /// <summary>
    /// 加粗
    /// </summary>
    public bool IsBold
    {
        get => Style.IsBold;
        set => Style.IsBold = value;
    }

    /// <summary>
    /// 自动换行
    /// </summary>
    public bool WrapText
    {
        get => Style.WrapText;
        set => Style.WrapText = value;
    }

    /// <summary>
    /// 字体颜色
    /// </summary>
    public short FontColor
    {
        get => Style.FontColor;
        set => Style.FontColor = value;
    }

    /// <summary>
    /// 字体大小
    /// </summary>
    public int FontSize
    {
        get => Style.FontSize;
        set => Style.FontSize = value;
    }

    /// <summary>
    /// 字体名称
    /// </summary>
    public string FontName
    {
        get => Style.FontName;
        set => Style.FontName = value;
    }

    /// <summary>
    /// 前景色
    /// </summary>
    public short FillForegroundColor
    {
        get => Style.FillForegroundColor;
        set => Style.FillForegroundColor = value;
    }

    /// <summary>
    /// 水平对齐方式
    /// </summary>
    public HorizontalAlignment HorizontalAlignment
    {
        get => Style.HorizontalAlignment;
        set => Style.HorizontalAlignment = value;
    }

    /// <summary>
    /// 垂直对齐方式
    /// </summary>
    public VerticalAlignment VerticalAlignment
    {
        get => Style.VerticalAlignment;
        set => Style.VerticalAlignment = value;
    }

    /// <summary>
    /// 初始化一个<see cref="StyleAttribute"/>类型的实例
    /// </summary>
    public StyleAttribute() => Style = new CellStyle();
}
