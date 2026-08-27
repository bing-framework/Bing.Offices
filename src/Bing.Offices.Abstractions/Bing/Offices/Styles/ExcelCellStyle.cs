using System;
using System.Collections.Generic;

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

/// <summary>
/// 单元格填充模式。
/// </summary>
public enum ExcelFillPattern
{
    None,
    Solid,
    LightGray,
    DarkGray
}

/// <summary>
/// 边框线型。
/// </summary>
public enum ExcelBorderLineStyle
{
    None,
    Thin,
    Medium,
    Thick,
    Dashed,
    Dotted,
    Double
}

/// <summary>
/// 水平对齐方式。
/// </summary>
public enum ExcelHorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify
}

/// <summary>
/// 垂直对齐方式。
/// </summary>
public enum ExcelVerticalAlignment
{
    Bottom,
    Center,
    Top,
    Justify
}

/// <summary>
/// 单个边框的提供程序无关描述。
/// </summary>
public sealed class ExcelBorderStyle
{
    /// <summary>
    /// 获取或设置线型。
    /// </summary>
    public ExcelBorderLineStyle LineStyle { get; init; }

    /// <summary>
    /// 获取或设置线条颜色。
    /// </summary>
    public ExcelColor Color { get; init; }
}

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

/// <summary>
/// 提供程序无关的单元格样式。
/// </summary>
public sealed class ExcelCellStyle
{
    /// <summary>
    /// 获取或设置字体名称。
    /// </summary>
    public string FontName { get; init; }

    /// <summary>
    /// 获取或设置字号。
    /// </summary>
    public short? FontSize { get; init; }

    /// <summary>
    /// 获取或设置粗体。
    /// </summary>
    public bool? Bold { get; init; }

    /// <summary>
    /// 获取或设置斜体。
    /// </summary>
    public bool? Italic { get; init; }

    /// <summary>
    /// 获取或设置下划线。
    /// </summary>
    public bool? Underline { get; init; }

    /// <summary>
    /// 获取或设置字体颜色。
    /// </summary>
    public ExcelColor FontColor { get; init; }

    /// <summary>
    /// 获取或设置前景色。
    /// </summary>
    public ExcelColor ForegroundColor { get; init; }

    /// <summary>
    /// 获取或设置背景色。
    /// </summary>
    public ExcelColor BackgroundColor { get; init; }

    /// <summary>
    /// 获取或设置填充模式。
    /// </summary>
    public ExcelFillPattern FillPattern { get; init; }

    /// <summary>
    /// 获取或设置上边框。
    /// </summary>
    public ExcelBorderStyle TopBorder { get; init; }

    /// <summary>
    /// 获取或设置下边框。
    /// </summary>
    public ExcelBorderStyle BottomBorder { get; init; }

    /// <summary>
    /// 获取或设置左边框。
    /// </summary>
    public ExcelBorderStyle LeftBorder { get; init; }

    /// <summary>
    /// 获取或设置右边框。
    /// </summary>
    public ExcelBorderStyle RightBorder { get; init; }

    /// <summary>
    /// 获取或设置水平对齐方式。
    /// </summary>
    public ExcelHorizontalAlignment HorizontalAlignment { get; init; }

    /// <summary>
    /// 获取或设置垂直对齐方式。
    /// </summary>
    public ExcelVerticalAlignment VerticalAlignment { get; init; }

    /// <summary>
    /// 获取或设置是否自动换行。
    /// </summary>
    public bool? WrapText { get; init; }

    /// <summary>
    /// 获取或设置缩进量。
    /// </summary>
    public short? Indent { get; init; }

    /// <summary>
    /// 获取或设置数字格式。
    /// </summary>
    public string NumberFormat { get; init; }

    /// <summary>
    /// 获取或设置显式清除/恢复默认描述；null 表示不清除任何属性。
    /// </summary>
    public ExcelCellStyleReset Reset { get; init; }
}
