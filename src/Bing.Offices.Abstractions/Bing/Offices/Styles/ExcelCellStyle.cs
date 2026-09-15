namespace Bing.Offices.Styles;

/// <summary>与 Excel 提供程序无关的单元格样式快照。</summary>
public sealed class ExcelCellStyle
{
    /// <summary>
    /// 获取或初始化字体名称。
    /// </summary>
    public string FontName { get; init; }

    /// <summary>
    /// 获取或初始化字号。
    /// </summary>
    public short? FontSize { get; init; }

    /// <summary>
    /// 获取或初始化粗体。
    /// </summary>
    public bool? Bold { get; init; }

    /// <summary>
    /// 获取或初始化斜体。
    /// </summary>
    public bool? Italic { get; init; }

    /// <summary>
    /// 获取或初始化下划线。
    /// </summary>
    public bool? Underline { get; init; }

    /// <summary>
    /// 获取或初始化字体颜色。
    /// </summary>
    public ExcelColor FontColor { get; init; }

    /// <summary>
    /// 获取或初始化前景色。
    /// </summary>
    public ExcelColor ForegroundColor { get; init; }

    /// <summary>
    /// 获取或初始化背景色。
    /// </summary>
    public ExcelColor BackgroundColor { get; init; }

    /// <summary>
    /// 获取或初始化填充模式。
    /// </summary>
    public ExcelFillPattern FillPattern { get; init; }

    /// <summary>
    /// 获取或初始化上边框。
    /// </summary>
    public ExcelBorderStyle TopBorder { get; init; }

    /// <summary>
    /// 获取或初始化下边框。
    /// </summary>
    public ExcelBorderStyle BottomBorder { get; init; }

    /// <summary>
    /// 获取或初始化左边框。
    /// </summary>
    public ExcelBorderStyle LeftBorder { get; init; }

    /// <summary>
    /// 获取或初始化右边框。
    /// </summary>
    public ExcelBorderStyle RightBorder { get; init; }

    /// <summary>
    /// 获取或初始化水平对齐方式。
    /// </summary>
    public ExcelHorizontalAlignment HorizontalAlignment { get; init; }

    /// <summary>
    /// 获取或初始化垂直对齐方式。
    /// </summary>
    public ExcelVerticalAlignment VerticalAlignment { get; init; }

    /// <summary>
    /// 获取或初始化是否自动换行。
    /// </summary>
    public bool? WrapText { get; init; }

    /// <summary>
    /// 获取或初始化缩进量。
    /// </summary>
    public short? Indent { get; init; }

    /// <summary>
    /// 获取或初始化数字格式。
    /// </summary>
    public string NumberFormat { get; init; }

    /// <summary>
    /// 获取或初始化显式清除/恢复默认描述；null 表示不清除任何属性。
    /// </summary>
    public ExcelCellStyleReset Reset { get; init; }
}
