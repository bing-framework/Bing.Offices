using Bing.Offices.Enums.Styles;

namespace Bing.Offices.Exports;

/// <summary>
/// 基础样式
/// </summary>
public interface IBaseStyle
{
    /// <summary>
    /// 加粗
    /// </summary>
    bool IsBold { get; set; }

    /// <summary>
    /// 自动换行
    /// </summary>
    bool WrapText { get; set; }

    /// <summary>
    /// 字体颜色
    /// </summary>
    short FontColor { get; set; }

    /// <summary>
    /// 字体大小
    /// </summary>
    int FontSize { get; set; }

    /// <summary>
    /// 字体名称
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 前景色
    /// </summary>
    short FillForegroundColor { get; set; }

    /// <summary>
    /// 水平对齐方式
    /// </summary>
    HorizontalAlignment HorizontalAlignment { get; set; }

    /// <summary>
    /// 垂直对齐方式
    /// </summary>
    VerticalAlignment VerticalAlignment { get; set; }
}
