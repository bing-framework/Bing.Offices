namespace Bing.Offices.Styles;

/// <summary>
/// 单元格样式属性的显式清除/恢复默认描述。
/// </summary>
public sealed class ExcelCellStyleReset
{
    /// <summary>是否恢复字体名称默认值。</summary>
    public bool FontName { get; init; }

    /// <summary>是否恢复字号默认值。</summary>
    public bool FontSize { get; init; }

    /// <summary>是否恢复粗体默认值。</summary>
    public bool Bold { get; init; }

    /// <summary>是否恢复斜体默认值。</summary>
    public bool Italic { get; init; }

    /// <summary>是否恢复下划线默认值。</summary>
    public bool Underline { get; init; }

    /// <summary>是否清除字体颜色。</summary>
    public bool FontColor { get; init; }

    /// <summary>是否清除填充前景色。</summary>
    public bool ForegroundColor { get; init; }

    /// <summary>是否清除填充背景色。</summary>
    public bool BackgroundColor { get; init; }

    /// <summary>是否恢复填充模式默认值。</summary>
    public bool FillPattern { get; init; }

    /// <summary>是否清除上边框。</summary>
    public bool TopBorder { get; init; }

    /// <summary>是否清除下边框。</summary>
    public bool BottomBorder { get; init; }

    /// <summary>是否清除左边框。</summary>
    public bool LeftBorder { get; init; }

    /// <summary>是否清除右边框。</summary>
    public bool RightBorder { get; init; }

    /// <summary>是否恢复水平对齐默认值。</summary>
    public bool HorizontalAlignment { get; init; }

    /// <summary>是否恢复垂直对齐默认值。</summary>
    public bool VerticalAlignment { get; init; }

    /// <summary>是否恢复自动换行默认值。</summary>
    public bool WrapText { get; init; }

    /// <summary>是否恢复缩进默认值。</summary>
    public bool Indent { get; init; }

    /// <summary>是否清除数字格式。</summary>
    public bool NumberFormat { get; init; }
}

