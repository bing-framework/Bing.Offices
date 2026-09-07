using System;
using System.Collections.Generic;

namespace Bing.Offices.Styles;

/// <summary>
/// 单元格填充模式。
/// </summary>
public enum ExcelFillPattern
{
    /// <summary>无填充。</summary>
    None,
    /// <summary>实心填充。</summary>
    Solid,
    /// <summary>浅灰填充。</summary>
    LightGray,
    /// <summary>深灰填充。</summary>
    DarkGray
}

/// <summary>
/// 边框线型。
/// </summary>
public enum ExcelBorderLineStyle
{
    /// <summary>无边框。</summary>
    None,
    /// <summary>细线。</summary>
    Thin,
    /// <summary>中线。</summary>
    Medium,
    /// <summary>粗线。</summary>
    Thick,
    /// <summary>虚线。</summary>
    Dashed,
    /// <summary>点线。</summary>
    Dotted,
    /// <summary>双线。</summary>
    Double
}

/// <summary>
/// 水平对齐方式。
/// </summary>
public enum ExcelHorizontalAlignment
{
    /// <summary>常规对齐。</summary>
    General,
    /// <summary>左对齐。</summary>
    Left,
    /// <summary>居中。</summary>
    Center,
    /// <summary>右对齐。</summary>
    Right,
    /// <summary>填充。</summary>
    Fill,
    /// <summary>两端对齐。</summary>
    Justify
}

/// <summary>
/// 垂直对齐方式。
/// </summary>
public enum ExcelVerticalAlignment
{
    /// <summary>底部对齐。</summary>
    Bottom,
    /// <summary>居中对齐。</summary>
    Center,
    /// <summary>顶部对齐。</summary>
    Top,
    /// <summary>两端对齐。</summary>
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

