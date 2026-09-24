using Bing.Offices.Styles;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 将 provider-neutral 样式映射到 ClosedXML 单元格样式。
/// </summary>
internal static class ClosedXmlStyleAdapter
{
    /// <summary>
    /// 将公共单元格样式按未设置属性保留模板值的规则应用到单元格。
    /// </summary>
    /// <param name="cell">待设置样式的单元格。</param>
    /// <param name="style">公共单元格样式。</param>
    public static void Apply(IXLCell cell, ExcelCellStyle style)
    {
        if (cell == null || style == null)
            return;
        ApplyReset(cell, style.Reset);
        if (!string.IsNullOrWhiteSpace(style.FontName))
            cell.Style.Font.FontName = style.FontName;
        if (style.FontSize.HasValue)
            cell.Style.Font.FontSize = style.FontSize.Value;
        if (style.Bold.HasValue)
            cell.Style.Font.Bold = style.Bold.Value;
        if (style.Italic.HasValue)
            cell.Style.Font.Italic = style.Italic.Value;
        if (style.Underline.HasValue)
            cell.Style.Font.Underline = style.Underline.Value ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
        ApplyColor(style.FontColor, color => cell.Style.Font.FontColor = color);
        ApplyColor(style.ForegroundColor, color => cell.Style.Fill.BackgroundColor = color);
        ApplyColor(style.BackgroundColor, color => cell.Style.Fill.PatternColor = color);
        if (style.FillPattern != ExcelFillPattern.None)
            cell.Style.Fill.PatternType = style.FillPattern == ExcelFillPattern.Solid
                ? XLFillPatternValues.Solid : XLFillPatternValues.DarkGray;
        ApplyBorder(cell.Style.Border, style.TopBorder, BorderSide.Top);
        ApplyBorder(cell.Style.Border, style.BottomBorder, BorderSide.Bottom);
        ApplyBorder(cell.Style.Border, style.LeftBorder, BorderSide.Left);
        ApplyBorder(cell.Style.Border, style.RightBorder, BorderSide.Right);
        if (style.WrapText.HasValue)
            cell.Style.Alignment.WrapText = style.WrapText.Value;
        if (style.Indent.HasValue)
            cell.Style.Alignment.Indent = style.Indent.Value;
        // General/Bottom are the public model defaults. Leaving them untouched
        // preserves a template's explicit alignment when an overlay only sets
        // font, fill or number format.
        if (style.HorizontalAlignment != ExcelHorizontalAlignment.General)
            cell.Style.Alignment.Horizontal = ToHorizontal(style.HorizontalAlignment);
        if (style.VerticalAlignment != ExcelVerticalAlignment.Bottom)
            cell.Style.Alignment.Vertical = ToVertical(style.VerticalAlignment);
        if (!string.IsNullOrWhiteSpace(style.NumberFormat))
            cell.Style.NumberFormat.Format = style.NumberFormat;
    }

    /// <summary>
    /// 按公共重置配置恢复单元格的默认样式属性。
    /// </summary>
    /// <param name="cell">待重置样式的单元格。</param>
    /// <param name="reset">样式重置配置。</param>
    private static void ApplyReset(IXLCell cell, ExcelCellStyleReset reset)
    {
        if (reset == null)
            return;
        if (reset.FontName)
            cell.Style.Font.FontName = "Calibri";
        if (reset.FontSize)
            cell.Style.Font.FontSize = 11;
        if (reset.Bold)
            cell.Style.Font.Bold = false;
        if (reset.Italic)
            cell.Style.Font.Italic = false;
        if (reset.Underline)
            cell.Style.Font.Underline = XLFontUnderlineValues.None;
        if (reset.FontColor)
            cell.Style.Font.FontColor = XLColor.Black;
        if (reset.ForegroundColor)
            cell.Style.Fill.BackgroundColor = XLColor.NoColor;
        if (reset.BackgroundColor)
            cell.Style.Fill.PatternColor = XLColor.NoColor;
        if (reset.FillPattern)
            cell.Style.Fill.PatternType = XLFillPatternValues.None;
        if (reset.TopBorder)
        {
            cell.Style.Border.TopBorder = XLBorderStyleValues.None;
            cell.Style.Border.TopBorderColor = XLColor.NoColor;
        }
        if (reset.BottomBorder)
        {
            cell.Style.Border.BottomBorder = XLBorderStyleValues.None;
            cell.Style.Border.BottomBorderColor = XLColor.NoColor;
        }
        if (reset.LeftBorder)
        {
            cell.Style.Border.LeftBorder = XLBorderStyleValues.None;
            cell.Style.Border.LeftBorderColor = XLColor.NoColor;
        }
        if (reset.RightBorder)
        {
            cell.Style.Border.RightBorder = XLBorderStyleValues.None;
            cell.Style.Border.RightBorderColor = XLColor.NoColor;
        }
        if (reset.HorizontalAlignment)
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
        if (reset.VerticalAlignment)
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
        if (reset.WrapText)
            cell.Style.Alignment.WrapText = false;
        if (reset.Indent)
            cell.Style.Alignment.Indent = 0;
        if (reset.NumberFormat)
            cell.Style.NumberFormat.Format = "General";
    }

    /// <summary>
    /// 设置单元格数字格式。
    /// </summary>
    /// <param name="cell">待设置的单元格。</param>
    /// <param name="format">数字格式字符串。</param>
    public static void ApplyNumberFormat(IXLCell cell, string format)
    {
        if (cell != null && !string.IsNullOrWhiteSpace(format))
            cell.Style.NumberFormat.Format = format;
    }

    /// <summary>
    /// 将公共颜色应用到 ClosedXML 样式属性。
    /// </summary>
    /// <param name="color">公共颜色。</param>
    /// <param name="setter">ClosedXML 颜色设置器。</param>
    private static void ApplyColor(ExcelColor color, Action<XLColor> setter)
    {
        if (color == null || string.IsNullOrWhiteSpace(color.Argb))
            return;
        var value = color.Argb.Trim().TrimStart('#');
        if (value.Length == 6)
            value = "FF" + value;
        setter(XLColor.FromHtml("#" + value.Substring(Math.Max(0, value.Length - 6))));
    }

    /// <summary>
    /// 将公共边框样式应用到指定边。
    /// </summary>
    /// <param name="border">ClosedXML 边框对象。</param>
    /// <param name="style">公共边框样式。</param>
    /// <param name="side">目标边。</param>
    private static void ApplyBorder(IXLBorder border, ExcelBorderStyle style, BorderSide side)
    {
        if (border == null || style == null)
            return;
        var line = ToBorder(style.LineStyle);
        switch (side)
        {
            case BorderSide.Top:
                border.TopBorder = line;
                ApplyColor(style.Color, color => border.TopBorderColor = color);
                break;
            case BorderSide.Bottom:
                border.BottomBorder = line;
                ApplyColor(style.Color, color => border.BottomBorderColor = color);
                break;
            case BorderSide.Left:
                border.LeftBorder = line;
                ApplyColor(style.Color, color => border.LeftBorderColor = color);
                break;
            case BorderSide.Right:
                border.RightBorder = line;
                ApplyColor(style.Color, color => border.RightBorderColor = color);
                break;
        }
    }

    /// <summary>
    /// 将公共边框线型转换为 ClosedXML 线型。
    /// </summary>
    /// <param name="value">公共边框线型。</param>
    /// <returns>对应的 ClosedXML 线型。</returns>
    private static XLBorderStyleValues ToBorder(ExcelBorderLineStyle value) => value switch
    {
        ExcelBorderLineStyle.Thin => XLBorderStyleValues.Thin,
        ExcelBorderLineStyle.Medium => XLBorderStyleValues.Medium,
        ExcelBorderLineStyle.Thick => XLBorderStyleValues.Thick,
        ExcelBorderLineStyle.Dashed => XLBorderStyleValues.Dashed,
        ExcelBorderLineStyle.Dotted => XLBorderStyleValues.Dotted,
        ExcelBorderLineStyle.Double => XLBorderStyleValues.Double,
        _ => XLBorderStyleValues.None
    };

    /// <summary>
    /// 表示待设置的边框方向。
    /// </summary>
    private enum BorderSide
    {
        /// <summary>
        /// 上边框。
        /// </summary>
        Top,
        /// <summary>
        /// 下边框。
        /// </summary>
        Bottom,
        /// <summary>
        /// 左边框。
        /// </summary>
        Left,
        /// <summary>
        /// 右边框。
        /// </summary>
        Right
    }

    /// <summary>
    /// 将公共水平对齐方式转换为 ClosedXML 对齐方式。
    /// </summary>
    /// <param name="value">公共水平对齐方式。</param>
    /// <returns>对应的 ClosedXML 水平对齐方式。</returns>
    private static XLAlignmentHorizontalValues ToHorizontal(ExcelHorizontalAlignment value) => value switch
    {
        ExcelHorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
        ExcelHorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
        ExcelHorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
        ExcelHorizontalAlignment.Fill => XLAlignmentHorizontalValues.Fill,
        ExcelHorizontalAlignment.Justify => XLAlignmentHorizontalValues.Justify,
        _ => XLAlignmentHorizontalValues.General
    };

    /// <summary>
    /// 将公共垂直对齐方式转换为 ClosedXML 对齐方式。
    /// </summary>
    /// <param name="value">公共垂直对齐方式。</param>
    /// <returns>对应的 ClosedXML 垂直对齐方式。</returns>
    private static XLAlignmentVerticalValues ToVertical(ExcelVerticalAlignment value) => value switch
    {
        ExcelVerticalAlignment.Top => XLAlignmentVerticalValues.Top,
        ExcelVerticalAlignment.Center => XLAlignmentVerticalValues.Center,
        ExcelVerticalAlignment.Justify => XLAlignmentVerticalValues.Justify,
        _ => XLAlignmentVerticalValues.Bottom
    };
}
