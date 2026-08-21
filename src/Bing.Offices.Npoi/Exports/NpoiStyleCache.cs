using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
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

    private sealed class Cache
    {
        private readonly Dictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>(StringComparer.Ordinal);
        private readonly Dictionary<string, ICellStyle> _composedStyles = new Dictionary<string, ICellStyle>(StringComparer.Ordinal);
        private readonly Dictionary<string, IFont> _fonts = new Dictionary<string, IFont>(StringComparer.Ordinal);

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

        private void ApplyStyle(IWorkbook workbook, ICellStyle style, ExcelCellStyle definition)
        {
            ApplyStyle(workbook, style, definition, false);
        }

        private void ApplyStyle(IWorkbook workbook, ICellStyle style, ExcelCellStyle definition, bool overlay)
        {
            var hasFontDefinition = !string.IsNullOrWhiteSpace(definition.FontName)
                || definition.FontSize.HasValue || definition.Bold.HasValue || definition.Italic.HasValue
                || definition.Underline.HasValue || definition.FontColor != null;
            if (!overlay || hasFontDefinition)
            {
                var fontKey = string.Join("|", definition.FontName, definition.FontSize, definition.Bold,
                    definition.Italic, definition.Underline, definition.FontColor?.Argb,
                    overlay ? style.FontIndex.ToString() : string.Empty);
                lock (_fonts)
                {
                    if (!_fonts.TryGetValue(fontKey, out var font))
                    {
                        font = workbook.CreateFont();
                        if (overlay)
                            font.CloneStyleFrom(workbook.GetFontAt(style.FontIndex));
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
                        ApplyFontColor(workbook, font, definition.FontColor);
                        _fonts.Add(fontKey, font);
                    }
                    style.SetFont(font);
                }
            }

            if (definition.FillPattern != ExcelFillPattern.None)
            {
                style.FillPattern = ToFillPattern(definition.FillPattern);
                ApplyFillColor(workbook, style, definition.ForegroundColor ?? definition.BackgroundColor);
            }
            ApplyBorder(workbook, style, BorderSide.Top, definition.TopBorder);
            ApplyBorder(workbook, style, BorderSide.Bottom, definition.BottomBorder);
            ApplyBorder(workbook, style, BorderSide.Left, definition.LeftBorder);
            ApplyBorder(workbook, style, BorderSide.Right, definition.RightBorder);
            if (!overlay || definition.HorizontalAlignment != ExcelHorizontalAlignment.General)
                style.Alignment = ToHorizontalAlignment(definition.HorizontalAlignment);
            if (!overlay || definition.VerticalAlignment != ExcelVerticalAlignment.Bottom)
                style.VerticalAlignment = ToVerticalAlignment(definition.VerticalAlignment);
            if (definition.WrapText.HasValue)
                style.WrapText = definition.WrapText.Value;
            if (definition.Indent.HasValue)
                style.Indention = definition.Indent.Value;
            if (!string.IsNullOrWhiteSpace(definition.NumberFormat))
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

        private static void ApplyFillColor(IWorkbook workbook, ICellStyle style, ExcelColor color)
        {
            if (color == null)
                return;
            if (style is XSSFCellStyle xssfStyle)
            {
                xssfStyle.SetFillForegroundColor(new XSSFColor(ParseRgb(color.Argb), null));
                return;
            }
            style.FillForegroundColor = ResolveHssfColor(color.Argb);
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
            var color = ResolveHssfColor(border.Color?.Argb);
            switch (side)
            {
                case BorderSide.Top:
                    style.BorderTop = lineStyle;
                    style.TopBorderColor = color;
                    break;
                case BorderSide.Bottom:
                    style.BorderBottom = lineStyle;
                    style.BottomBorderColor = color;
                    break;
                case BorderSide.Left:
                    style.BorderLeft = lineStyle;
                    style.LeftBorderColor = color;
                    break;
                default:
                    style.BorderRight = lineStyle;
                    style.RightBorderColor = color;
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
            style.Indent, style.NumberFormat);
    }

    private enum BorderSide
    {
        Top,
        Bottom,
        Left,
        Right
    }
}
