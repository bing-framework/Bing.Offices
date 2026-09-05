using System;
using System.Globalization;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;

namespace Bing.Offices.Dates;

/// <summary>无 offset 文本转换为 DateTimeOffset 时使用的策略。</summary>
public enum ExcelDateOffsetPolicy
{
    /// <summary>必须由输入文本显式提供 offset。</summary>
    RequireExplicitOffset,
    /// <summary>使用配置的固定 offset。</summary>
    UseFixedOffset
}

/// <summary>Excel 与 CSV 共用的确定性日期解析器。</summary>
internal static class ExcelDateParser
{
    private const string DefaultDateFormat = "yyyy-MM-dd";

    /// <summary>
    /// 按日期特性配置将单元格值转换为目标日期类型。
    /// </summary>
    /// <param name="cell">原始单元格描述。</param>
    /// <param name="text">规范化文本。</param>
    /// <param name="targetType">目标类型。</param>
    /// <param name="culture">请求区域性。</param>
    /// <param name="attribute">可选日期特性。</param>
    /// <param name="value">转换结果。</param>
    /// <returns>能够转换时为 true。</returns>
    public static bool TryParse(ExcelCellValue cell, string text, Type targetType, CultureInfo culture,
        ExcelDateAttribute attribute, out object value)
    {
        value = null;
        var effectiveType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (effectiveType != typeof(DateTime) && effectiveType != typeof(DateTimeOffset))
            return false;
        if (cell?.Value is DateTime dateTime && effectiveType == typeof(DateTime))
        {
            value = DateTime.SpecifyKind(dateTime, DateTimeKind.Unspecified);
            return true;
        }
        if (cell?.Value is DateTimeOffset dateTimeOffset && effectiveType == typeof(DateTimeOffset))
        {
            value = dateTimeOffset;
            return true;
        }
        if (cell?.Value is double serial && effectiveType == typeof(DateTime))
        {
            try
            {
                value = DateTime.SpecifyKind(FromExcelSerial(serial, cell.IsDate1904), DateTimeKind.Unspecified);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
        var effectiveCulture = culture ?? CultureInfo.InvariantCulture;
        if (!string.IsNullOrWhiteSpace(attribute?.CultureName))
            effectiveCulture = CultureInfo.GetCultureInfo(attribute.CultureName);
        var format = attribute?.Format;
        if (effectiveType == typeof(DateTime))
        {
            if (DateTime.TryParseExact(text, format ?? DefaultDateFormat, effectiveCulture,
                    DateTimeStyles.None, out var parsed))
            {
                value = DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified);
                return true;
            }
            return false;
        }

        var offsetPolicy = attribute?.OffsetPolicy ?? ExcelDateOffsetPolicy.RequireExplicitOffset;
        var offsetMinutes = attribute?.OffsetMinutes;
        if (offsetPolicy == ExcelDateOffsetPolicy.UseFixedOffset && !offsetMinutes.HasValue)
            return false;
        if (HasExplicitOffset(text))
        {
            var formats = format == null
                ? new[] { "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK", "yyyy-MM-dd'T'HH:mm:ssK" }
                : new[] { format };
            foreach (var candidate in formats)
            {
                if (DateTimeOffset.TryParseExact(text, candidate, effectiveCulture,
                        DateTimeStyles.None, out var parsed))
                {
                    value = parsed;
                    return true;
                }
            }
            return false;
        }
        if (offsetPolicy != ExcelDateOffsetPolicy.UseFixedOffset)
            return false;
        var localFormat = format ?? DefaultDateFormat;
        if (!DateTime.TryParseExact(text, localFormat, effectiveCulture, DateTimeStyles.None,
                out var localDate))
            return false;
        try
        {
            value = new DateTimeOffset(DateTime.SpecifyKind(localDate, DateTimeKind.Unspecified),
                TimeSpan.FromMinutes(offsetMinutes.Value));
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// 按 Workbook Data Validation 的日期或时间语义解析值。
    /// </summary>
    /// <param name="cell">原始单元格值。</param>
    /// <param name="text">单元格或约束的文本值。</param>
    /// <param name="timeOnly">是否只解析时间部分。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="value">解析后的无时区日期时间。</param>
    /// <returns>解析成功时为 true。</returns>
    internal static bool TryParseValidation(ExcelCellValue cell, string text, bool timeOnly,
        bool isDate1904, out DateTime value)
    {
        value = default;
        if (cell?.Value is DateTime date)
        {
            value = DateTime.SpecifyKind(timeOnly ? DateTime.MinValue.Add(date.TimeOfDay) : date,
                DateTimeKind.Unspecified);
            return true;
        }
        if (cell?.Value is double serial)
            return TryFromExcelSerial(serial, timeOnly, isDate1904, out value);
        if (string.IsNullOrWhiteSpace(text))
            return false;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var textSerial))
            return TryFromExcelSerial(textSerial, timeOnly, isDate1904, out value);
        if (timeOnly)
        {
            if (!DateTime.TryParseExact(text, new[] { "HH:mm:ss.FFFFFFF", "HH:mm:ss", "H:mm:ss" },
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var time))
                return false;
            value = DateTime.MinValue.Add(time.TimeOfDay);
            return true;
        }
        if (!TryParse(new ExcelCellValue(text, text, ExcelCellKind.Text, isDate1904: isDate1904), text,
                typeof(DateTime), CultureInfo.InvariantCulture, null, out var parsed))
            return false;
        value = (DateTime)parsed;
        return true;
    }

    private static bool HasExplicitOffset(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;
        if (text.EndsWith("Z", StringComparison.OrdinalIgnoreCase))
            return true;
        var separator = text.LastIndexOfAny(new[] { '+', '-' });
        return separator > 10 && text.Length - separator >= 6
            && text[separator + 3] == ':';
    }

    private static DateTime FromExcelSerial(double serial, bool isDate1904)
    {
        if (isDate1904)
            return new DateTime(1904, 1, 1).AddDays(serial);
        var baseDate = new DateTime(1899, 12, 31);
        return baseDate.AddDays(serial >= 60 ? serial - 1 : serial);
    }

    private static bool TryFromExcelSerial(double serial, bool timeOnly, bool isDate1904,
        out DateTime value)
    {
        try
        {
            var date = FromExcelSerial(serial, isDate1904);
            value = DateTime.SpecifyKind(timeOnly ? DateTime.MinValue.Add(date.TimeOfDay) : date,
                DateTimeKind.Unspecified);
            return true;
        }
        catch (ArgumentException)
        {
            value = default;
            return false;
        }
    }
}
