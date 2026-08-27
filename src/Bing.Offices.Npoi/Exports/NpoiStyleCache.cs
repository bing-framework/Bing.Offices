using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Bing.Offices.Attributes;
using Bing.Offices.Styles;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// Workbook 级规范化样式缓存。
/// </summary>
internal static class NpoiStyleCache
{
    private static readonly ConditionalWeakTable<IWorkbook, Cache> Caches = new ConditionalWeakTable<IWorkbook, Cache>();

    /// <summary>
    /// 获取或创建规范化样式。
    /// </summary>
    internal static ICellStyle GetOrCreate(IWorkbook workbook, ExcelCellStyle definition)
    {
        if (definition == null)
            return workbook.GetCellStyleAt(0);
        return Caches.GetValue(workbook, _ => new Cache()).GetOrCreate(workbook, definition);
    }

    /// <summary>
    /// 将请求样式逐属性叠加到现有样式，保留模板中的未覆盖属性。
    /// </summary>
    internal static ICellStyle Compose(IWorkbook workbook, ICellStyle baseStyle, ExcelCellStyle overlay)
    {
        if (overlay == null)
            return baseStyle;
        return Caches.GetValue(workbook, _ => new Cache()).Compose(workbook, baseStyle, overlay);
    }

    /// <summary>
    /// 使用 Workbook 级缓存应用实体表头字体，避免按单元格重复创建样式和字体。
    /// </summary>
    internal static ICellStyle ApplyHeaderAttribute(IWorkbook workbook, ICellStyle baseStyle,
        HeaderAttribute attribute)
    {
        return Caches.GetValue(workbook, _ => new Cache())
            .ApplyHeaderAttribute(workbook, baseStyle, attribute);
    }

    private sealed class Cache
    {
        private readonly Dictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>(StringComparer.Ordinal);
        private readonly Dictionary<string, ICellStyle> _composedStyles = new Dictionary<string, ICellStyle>(StringComparer.Ordinal);
        private readonly Dictionary<string, IFont> _fonts = new Dictionary<string, IFont>(StringComparer.Ordinal);
        private readonly Dictionary<string, ICellStyle> _headerStyles = new Dictionary<string, ICellStyle>(StringComparer.Ordinal);

        internal ICellStyle GetOrCreate(IWorkbook workbook, ExcelCellStyle definition)
        {
            var key = CreateKey(definition);
            lock (_styles)
            {
                if (_styles.TryGetValue(key, out var style))
                    return style;
                style = workbook.CreateCellStyle();
                ApplyStyle(workbook, style, definition);
                _styles.Add(key, style);
                return style;
            }
        }

        internal ICellStyle Compose(IWorkbook workbook, ICellStyle baseStyle, ExcelCellStyle definition)
        {
            var key = baseStyle.Index + ":" + CreateKey(definition);
            lock (_composedStyles)
            {
                if (_composedStyles.TryGetValue(key, out var style))
                    return style;
                style = workbook.CreateCellStyle();
                style.CloneStyleFrom(baseStyle);
                ApplyStyle(workbook, style, definition, true);
                _composedStyles.Add(key, style);
                return style;
            }
        }

        internal ICellStyle ApplyHeaderAttribute(IWorkbook workbook, ICellStyle baseStyle,
            HeaderAttribute attribute)
        {
            var fontKey = string.Join("|", baseStyle.FontIndex, attribute.FontName, attribute.FontSize,
                attribute.Bold, attribute.Color);
            IFont font;
            lock (_fonts)
            {
                if (!_fonts.TryGetValue(fontKey, out font))
                {
                    font = workbook.CreateFont();
                    font.CloneStyleFrom(workbook.GetFontAt(baseStyle.FontIndex));
                    font.FontName = attribute.FontName;
                    font.Color = Resolvers.ColorResolver.Resolve(attribute.Color);
                    font.FontHeightInPoints = attribute.FontSize;
                    font.IsBold = attribute.Bold;
                    _fonts.Add(fontKey, font);
                }
            }

            var styleKey = string.Join("|", baseStyle.Index, fontKey);
            lock (_headerStyles)
            {
                if (_headerStyles.TryGetValue(styleKey, out var style))
                    return style;
                style = workbook.CreateCellStyle();
                style.CloneStyleFrom(baseStyle);
                style.SetFont(font);
                _headerStyles.Add(styleKey, style);
                return style;
            }
        }

        private void ApplyStyle(IWorkbook workbook, ICellStyle style, ExcelCellStyle definition)
        {
            ApplyStyle(workbook, style, definition, false);
        }

        private void ApplyStyle(IWorkbook workbook, ICellStyle style, ExcelCellStyle definition, bool overlay)
        {
            var reset = definition.Reset;
            var hasFontDefinition = !string.IsNullOrWhiteSpace(definition.FontName)
                || definition.FontSize.HasValue || definition.Bold.HasValue || definition.Italic.HasValue
                || definition.Underline.HasValue || definition.FontColor != null
                || reset?.FontName == true || reset?.FontSize == true || reset?.Bold == true
                || reset?.Italic == true || reset?.Underline == true || reset?.FontColor == true;
            if (!overlay || hasFontDefinition)
            {
                var fontKey = string.Join("|", definition.FontName, definition.FontSize, definition.Bold,
                    definition.Italic, definition.Underline, definition.FontColor?.Argb,
                    reset?.FontName, reset?.FontSize, reset?.Bold, reset?.Italic, reset?.Underline,
                    reset?.FontColor,
                    overlay ? style.FontIndex.ToString() : string.Empty);
                lock (_fonts)
                {
                    if (!_fonts.TryGetValue(fontKey, out var font))
                    {
                        font = workbook.CreateFont();
                        if (overlay)
                            font.CloneStyleFrom(workbook.GetFontAt(style.FontIndex));
                        var defaultFont = workbook.GetFontAt(0);
                        if (reset?.FontName == true)
                            font.FontName = defaultFont.FontName;
                        if (reset?.FontSize == true)
                            font.FontHeightInPoints = defaultFont.FontHeightInPoints;
                        if (reset?.Bold == true)
                            font.IsBold = defaultFont.IsBold;
                        if (reset?.Italic == true)
                            font.IsItalic = defaultFont.IsItalic;
                        if (reset?.Underline == true)
                            font.Underline = defaultFont.Underline;
                        if (!string.IsNullOrWhiteSpace(definition.FontName))
                            font.FontName = definition.FontName;
                        if (definition.FontSize.HasValue)
                            font.FontHeightInPoints = definition.FontSize.Value;
                        if (definition.Bold.HasValue)
                            font.IsBold = definition.Bold.Value;
                        if (definition.Italic.HasValue)
                            font.IsItalic = definition.Italic.Value;
                        if (definition.Underline.HasValue)
                            font.Underline = definition.Underline.Value
                                ? FontUnderlineType.Single
                                : FontUnderlineType.None;
                        if (reset?.FontColor == true)
                            font.Color = defaultFont.Color;
                        else
                            ApplyFontColor(workbook, font, definition.FontColor);
                        _fonts.Add(fontKey, font);
                    }
                    style.SetFont(font);
                }
            }

            if (!overlay || definition.FillPattern != ExcelFillPattern.None || reset?.FillPattern == true)
            {
                style.FillPattern = reset?.FillPattern == true ? FillPattern.NoFill : ToFillPattern(definition.FillPattern);
            }
            if (reset?.ForegroundColor == true)
                ApplyFillColor(workbook, style, null, true);
            else if (definition.ForegroundColor != null)
                ApplyFillColor(workbook, style, definition.ForegroundColor, true);
            if (reset?.BackgroundColor == true)
                ApplyFillColor(workbook, style, null, false);
            else if (definition.BackgroundColor != null)
                ApplyFillColor(workbook, style, definition.BackgroundColor, false);
            if (reset?.TopBorder == true)
                style.BorderTop = BorderStyle.None;
            if (reset?.BottomBorder == true)
                style.BorderBottom = BorderStyle.None;
            if (reset?.LeftBorder == true)
                style.BorderLeft = BorderStyle.None;
            if (reset?.RightBorder == true)
                style.BorderRight = BorderStyle.None;
            ApplyBorder(workbook, style, BorderSide.Top, definition.TopBorder);
            ApplyBorder(workbook, style, BorderSide.Bottom, definition.BottomBorder);
            ApplyBorder(workbook, style, BorderSide.Left, definition.LeftBorder);
            ApplyBorder(workbook, style, BorderSide.Right, definition.RightBorder);
            if (!overlay || definition.HorizontalAlignment != ExcelHorizontalAlignment.General
                || reset?.HorizontalAlignment == true)
                style.Alignment = ToHorizontalAlignment(definition.HorizontalAlignment);
            if (!overlay || definition.VerticalAlignment != ExcelVerticalAlignment.Bottom
                || reset?.VerticalAlignment == true)
                style.VerticalAlignment = ToVerticalAlignment(definition.VerticalAlignment);
            if (reset?.WrapText == true)
                style.WrapText = false;
            else if (definition.WrapText.HasValue)
                style.WrapText = definition.WrapText.Value;
            if (reset?.Indent == true)
                style.Indention = 0;
            else if (definition.Indent.HasValue)
                style.Indention = definition.Indent.Value;
            if (reset?.NumberFormat == true)
                style.DataFormat = workbook.CreateDataFormat().GetFormat("General");
            else if (!string.IsNullOrWhiteSpace(definition.NumberFormat))
                style.DataFormat = workbook.CreateDataFormat().GetFormat(definition.NumberFormat);
        }

        private static void ApplyFontColor(IWorkbook workbook, IFont font, ExcelColor color)
        {
            if (color == null)
                return;
            if (font is XSSFFont xssfFont)
            {
                xssfFont.SetColor(new XSSFColor(ParseRgb(color.Argb), null));
                return;
            }
            font.Color = ResolveHssfColor(color.Argb);
        }

        private static void ApplyFillColor(IWorkbook workbook, ICellStyle style, ExcelColor color, bool foreground)
        {
            if (style is XSSFCellStyle xssfStyle)
            {
                if (foreground)
                    xssfStyle.FillForegroundColorColor = color == null
                        ? null
                        : new XSSFColor(ParseRgb(color.Argb), null);
                else
                    xssfStyle.FillBackgroundColorColor = color == null
                        ? null
                        : new XSSFColor(ParseRgb(color.Argb), null);
                return;
            }
            var indexedColor = color == null ? IndexedColors.Automatic.Index : ResolveHssfColor(color.Argb);
            if (foreground)
                style.FillForegroundColor = indexedColor;
            else
                style.FillBackgroundColor = indexedColor;
        }

        private static byte[] ParseRgb(string value)
        {
            var argb = ParseArgb(value);
            return new[] { argb[1], argb[2], argb[3] };
        }

        private static void ApplyBorder(IWorkbook workbook, ICellStyle style, BorderSide side,
            ExcelBorderStyle border)
        {
            if (border == null)
                return;
            var lineStyle = ToBorderStyle(border.LineStyle);
            var indexedColor = style is XSSFCellStyle ? IndexedColors.Automatic.Index
                : ResolveHssfColor(border.Color?.Argb);
            switch (side)
            {
                case BorderSide.Top:
                    style.BorderTop = lineStyle;
                    SetBorderColor(style, side, border.Color, indexedColor);
                    break;
                case BorderSide.Bottom:
                    style.BorderBottom = lineStyle;
                    SetBorderColor(style, side, border.Color, indexedColor);
                    break;
                case BorderSide.Left:
                    style.BorderLeft = lineStyle;
                    SetBorderColor(style, side, border.Color, indexedColor);
                    break;
                default:
                    style.BorderRight = lineStyle;
                    SetBorderColor(style, side, border.Color, indexedColor);
                    break;
            }
        }

        /// <summary>
        /// 按 Workbook 格式写入边框颜色，XSSF 使用完整 RGB，HSSF 使用索引色板。
        /// </summary>
        /// <param name="style">目标单元格样式。</param>
        /// <param name="side">边框方向。</param>
        /// <param name="color">provider-neutral 颜色。</param>
        /// <param name="indexedColor">HSSF 索引颜色。</param>
        private static void SetBorderColor(ICellStyle style, BorderSide side, ExcelColor color, short indexedColor)
        {
            if (style is XSSFCellStyle xssfStyle && color != null)
            {
                var xssfColor = new XSSFColor(ParseRgb(color.Argb), null);
                switch (side)
                {
                    case BorderSide.Top:
                        xssfStyle.SetTopBorderColor(xssfColor);
                        break;
                    case BorderSide.Bottom:
                        xssfStyle.SetBottomBorderColor(xssfColor);
                        break;
                    case BorderSide.Left:
                        xssfStyle.SetLeftBorderColor(xssfColor);
                        break;
                    default:
                        xssfStyle.SetRightBorderColor(xssfColor);
                        break;
                }
                return;
            }
            switch (side)
            {
                case BorderSide.Top:
                    style.TopBorderColor = indexedColor;
                    break;
                case BorderSide.Bottom:
                    style.BottomBorderColor = indexedColor;
                    break;
                case BorderSide.Left:
                    style.LeftBorderColor = indexedColor;
                    break;
                default:
                    style.RightBorderColor = indexedColor;
                    break;
            }
        }

        private static byte[] ParseArgb(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return new byte[] { 0, 0, 0, 0 };
            var hex = value.TrimStart('#');
            if (hex.Length == 6)
                hex = "FF" + hex;
            if (hex.Length != 8)
                throw new ArgumentException("颜色必须是 6 或 8 位十六进制 ARGB。", nameof(value));
            var bytes = new byte[4];
            for (var index = 0; index < bytes.Length; index++)
                bytes[index] = Convert.ToByte(hex.Substring(index * 2, 2), 16);
            return bytes;
        }

        private static short ResolveHssfColor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return IndexedColors.Automatic.Index;
            var hex = value.TrimStart('#');
            if (hex.Length == 8)
                hex = hex.Substring(2);
            switch (hex.ToUpperInvariant())
            {
                case "000000": return IndexedColors.Black.Index;
                case "FFFFFF": return IndexedColors.White.Index;
                case "FF0000": return IndexedColors.Red.Index;
                case "00FF00": return IndexedColors.Green.Index;
                case "0000FF": return IndexedColors.Blue.Index;
                case "FFFF00": return IndexedColors.Yellow.Index;
                default: throw new NotSupportedException(".xls 不支持该自定义颜色: " + value);
            }
        }

        private static FillPattern ToFillPattern(ExcelFillPattern pattern) => pattern switch
        {
            ExcelFillPattern.Solid => FillPattern.SolidForeground,
            ExcelFillPattern.LightGray => FillPattern.LessDots,
            ExcelFillPattern.DarkGray => FillPattern.LeastDots,
            _ => FillPattern.NoFill
        };

        private static BorderStyle ToBorderStyle(ExcelBorderLineStyle style) => style switch
        {
            ExcelBorderLineStyle.Thin => BorderStyle.Thin,
            ExcelBorderLineStyle.Medium => BorderStyle.Medium,
            ExcelBorderLineStyle.Thick => BorderStyle.Thick,
            ExcelBorderLineStyle.Dashed => BorderStyle.Dashed,
            ExcelBorderLineStyle.Dotted => BorderStyle.Dotted,
            ExcelBorderLineStyle.Double => BorderStyle.Double,
            _ => BorderStyle.None
        };

        private static HorizontalAlignment ToHorizontalAlignment(ExcelHorizontalAlignment alignment) => alignment switch
        {
            ExcelHorizontalAlignment.Left => HorizontalAlignment.Left,
            ExcelHorizontalAlignment.Center => HorizontalAlignment.Center,
            ExcelHorizontalAlignment.Right => HorizontalAlignment.Right,
            ExcelHorizontalAlignment.Fill => HorizontalAlignment.Fill,
            ExcelHorizontalAlignment.Justify => HorizontalAlignment.Justify,
            _ => HorizontalAlignment.General
        };

        private static VerticalAlignment ToVerticalAlignment(ExcelVerticalAlignment alignment) => alignment switch
        {
            ExcelVerticalAlignment.Center => VerticalAlignment.Center,
            ExcelVerticalAlignment.Top => VerticalAlignment.Top,
            ExcelVerticalAlignment.Justify => VerticalAlignment.Justify,
            _ => VerticalAlignment.Bottom
        };

        private static string CreateKey(ExcelCellStyle style) => string.Join(";", style.FontName, style.FontSize,
            style.Bold, style.Italic, style.Underline, style.FontColor?.Argb, style.ForegroundColor?.Argb,
            style.BackgroundColor?.Argb, style.FillPattern, style.TopBorder?.LineStyle,
            style.TopBorder?.Color?.Argb, style.BottomBorder?.LineStyle, style.BottomBorder?.Color?.Argb,
            style.LeftBorder?.LineStyle, style.LeftBorder?.Color?.Argb, style.RightBorder?.LineStyle,
            style.RightBorder?.Color?.Argb, style.HorizontalAlignment, style.VerticalAlignment, style.WrapText,
            style.Indent, style.NumberFormat, style.Reset?.FontName, style.Reset?.FontSize,
            style.Reset?.Bold, style.Reset?.Italic, style.Reset?.Underline, style.Reset?.FontColor,
            style.Reset?.ForegroundColor, style.Reset?.BackgroundColor, style.Reset?.FillPattern,
            style.Reset?.TopBorder, style.Reset?.BottomBorder, style.Reset?.LeftBorder, style.Reset?.RightBorder,
            style.Reset?.HorizontalAlignment, style.Reset?.VerticalAlignment, style.Reset?.WrapText,
            style.Reset?.Indent, style.Reset?.NumberFormat);
    }

    private enum BorderSide
    {
        Top,
        Bottom,
        Left,
        Right
    }
}
