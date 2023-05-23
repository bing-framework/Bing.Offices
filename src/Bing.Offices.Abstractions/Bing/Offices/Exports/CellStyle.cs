using Bing.Offices.Enums.Styles;

namespace Bing.Offices.Exports;

/// <summary>
/// 单元格样式
/// </summary>
public class CellStyle : IBaseStyle
{
    /// <summary>
    /// 加粗
    /// </summary>
    public bool IsBold { get; set; } = false;

    /// <summary>
    /// 自动换行
    /// </summary>
    public bool WrapText { get; set; } = true;

    /// <summary>
    /// 字体颜色
    /// </summary>
    public short FontColor { get; set; } = DefaultStyle.FontColor;

    /// <summary>
    /// 字体大小
    /// </summary>
    public int FontSize { get; set; } = DefaultStyle.FontSize;

    /// <summary>
    /// 字体名称
    /// </summary>
    public string FontName { get; set; } = DefaultStyle.FontName;

    /// <summary>
    /// 前景色
    /// </summary>
    public short FillForegroundColor { get; set; }

    /// <summary>
    /// 水平对齐方式
    /// </summary>
    public HorizontalAlignment HorizontalAlignment { get; set; } = DefaultStyle.HorizontalAlignment;

    /// <summary>
    /// 垂直对齐方式
    /// </summary>
    public VerticalAlignment VerticalAlignment { get; set; } = DefaultStyle.VerticalAlignment;

    /// <summary>
    /// 输出字符串
    /// </summary>
    public override string ToString() => $"{IsBold}{WrapText}{FontColor}{FontSize}{FontName}{HorizontalAlignment}{VerticalAlignment}{FillForegroundColor}";
}
